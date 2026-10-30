"""Light update selection, fallback, and data-preserving installation."""
import hashlib
import json
import queue
import shutil
import subprocess
import zipfile
from pathlib import Path
from types import SimpleNamespace
import urllib.request
import pytest
import lp_tool
from test_v3_pool_gui import app
from test_v3_9_2_compat_update import package, collect_update


def make_light(path, runtime_id='a'*64, extra=None):
    with zipfile.ZipFile(path, 'w') as z:
        for name in ('lp_tool.py', 'lp_v3.py'):
            z.writestr('app/'+name, 'APP_VERSION = "9.9.9"\n')
        z.writestr('app/watermark_encrypt.py', '# test')
        z.writestr('app/start.bat', '@echo off')
        z.writestr('app/.update_manifest.json', json.dumps({'version':'9.9.9', 'runtime_id':runtime_id}))
        for name, data in (extra or {}).items(): z.writestr(name, data)
    return path


@pytest.mark.parametrize('scenario', ['compatible', 'mismatch', 'corrupt_marker', 'missing_marker', 'missing_python', 'missing_dependency', 'network'])
def test_light_update_and_full_fallback(app, tmp_path, monkeypatch, scenario):
    base=tmp_path/'app'; base.mkdir()
    runtime=base/'runtime'; runtime.mkdir()
    marker=runtime/'.docmailer-runtime-id'
    if scenario != 'missing_marker': marker.write_text('b'*64 if scenario == 'mismatch' else 'a'*64)
    if scenario == 'corrupt_marker': marker.write_bytes(b'\xff')
    if scenario != 'missing_python': (runtime/'python.exe').write_bytes(b'not run')
    (base/'lp_data').mkdir(); (base/'lp_data'/'settings.json').write_bytes(b'private')
    full=package(tmp_path/'full.zip'); light=make_light(tmp_path/'light.zip')
    calls=[]
    def retrieve(url, dest, reporthook=None):
        calls.append(url)
        if scenario == 'network' and url == 'light': raise OSError('offline')
        path=light if url=='light' else full
        shutil.copyfile(path,dest);reporthook(path.stat().st_size,1,path.stat().st_size)
    monkeypatch.setattr(urllib.request,'urlretrieve',retrieve)
    monkeypatch.setattr(subprocess,'run',lambda *a,**kw:SimpleNamespace(returncode=1 if scenario=='missing_dependency' else 0))
    monkeypatch.setattr(lp_tool,'app_dir',lambda:base)
    info={'tag':'v9.9.9','asset_url':'full','asset_size':full.stat().st_size,
          'asset_digest':'sha256:'+hashlib.sha256(full.read_bytes()).hexdigest(),
          'light_asset':{'asset_url':'light','asset_size':light.stat().st_size,
                         'asset_digest':'sha256:'+hashlib.sha256(light.read_bytes()).hexdigest()}}
    assert app._self_update(info)
    events=collect_update(app)
    assert any(k=='update_ready' for k,_ in events),events
    assert calls==(['light'] if scenario=='compatible' else ['full'] if scenario in ('missing_marker','missing_python') else ['light','full'])
    assert (base/'lp_data'/'settings.json').read_bytes()==b'private'
    assert not (base/'.update_pkg.zip').exists()
    assert not (base/'.update_staging'/'app'/'runtime').exists()


def test_light_digest_mismatch_fails_and_removes_temp(app,tmp_path,monkeypatch):
    base=tmp_path/'app';(base/'runtime').mkdir(parents=True)
    (base/'runtime'/'python.exe').write_bytes(b'x');(base/'runtime'/'.docmailer-runtime-id').write_text('a'*64)
    light=make_light(tmp_path/'light.zip')
    monkeypatch.setattr(lp_tool,'app_dir',lambda:base)
    monkeypatch.setattr(urllib.request,'urlretrieve',lambda url,dest,reporthook=None:shutil.copyfile(light,dest))
    app._self_update({'tag':'v9.9.9','asset_url':'full','light_asset':{'asset_url':'light','asset_digest':'sha256:bad'}})
    events=collect_update(app)
    assert any(k=='update_err' for k,_ in events)
    assert not (base/'.update_pkg.zip').exists() and not (base/'.update_staging').exists()


def test_preexisting_staging_user_data_is_not_deleted(app,tmp_path,monkeypatch):
    base=tmp_path/'app';data=base/'.update_staging'/'lp_data';data.mkdir(parents=True)
    (data/'private.json').write_bytes(b'keep')
    full=package(tmp_path/'full.zip')
    monkeypatch.setattr(lp_tool,'app_dir',lambda:base)
    monkeypatch.setattr(urllib.request,'urlretrieve',lambda url,dest,reporthook=None:shutil.copyfile(full,dest))
    app._self_update({'tag':'v9.9.9','asset_url':'full'})
    events=collect_update(app)
    assert any(k=='update_err' for k,_ in events)
    assert (data/'private.json').read_bytes()==b'keep'


def test_light_install_protects_runtime_data_and_removes_only_program_backup(tmp_path):
    base=tmp_path/'app'; source=base/'.update_staging'/'app';source.mkdir(parents=True)
    (source/'a.py').write_bytes(b'new');(source/'.update_manifest.json').write_text('{}')
    (base/'a.py').write_bytes(b'old')
    (base/'runtime').mkdir();(base/'runtime'/'library.dll').write_bytes(b'runtime')
    (base/'lp_data').mkdir();(base/'lp_data'/'events.json').write_bytes(b'user')
    helper=lp_tool._write_update_helper(base,source)
    run=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(helper),'-WaitProcessId','0','-NoUi','-NoRestart'],capture_output=True,text=True,timeout=30)
    result=json.loads((base/'.update_result.json').read_text(encoding='utf-8-sig'))
    assert run.returncode==0 and result['success']
    assert (base/'a.py').read_bytes()==b'new'
    assert (base/'runtime'/'library.dll').read_bytes()==b'runtime'
    assert (base/'lp_data'/'events.json').read_bytes()==b'user'
    assert not (base/'.update_manifest.json').exists() and not (base/'.update_staging').exists()
    assert not (Path(result['backup'])/'program').exists()
    assert (Path(result['backup'])/'lp_data'/'events.json').read_bytes()==b'user'
