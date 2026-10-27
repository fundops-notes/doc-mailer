"""Old-event upgrades and the actual update preparation/installer boundaries."""
import hashlib
import json
from pathlib import Path
import queue
import shutil
import subprocess
import threading
import time
import urllib.request
import zipfile

import pytest
import lp_tool
import lp_v3
from test_v3_pool_gui import app, no_messagebox
from test_v3_9_0_workflow import make_event, pdf, show, texts


def old_event(tmp_path, count=4):
    event, first = make_event(tmp_path)
    event['pipeline_ui_version'] = '3.9'
    event['config_version'] = '3.8.4'
    for index in range(1, count):
        item = lp_v3.new_item('pdf')
        item['params']['files'] = [pdf(tmp_path/('step%d.pdf' % index), 'STEP%d' % index)]
        event['pipeline'].append(item)
    previous = tmp_path/'it-20200101-010101-abcd12'
    previous.mkdir(exist_ok=True)
    path = pdf(previous/'old.pdf', 'OLD')
    event['process_overrides'] = {r['name']: [path] for r in event['recipients']}
    event['process_results'] = {r['name']: [path] for r in event['recipients']}
    return event


def write_raw(store, event):
    store.ensure_dirs()
    path = store.event_path(event['id'])
    path.write_text(json.dumps(event, ensure_ascii=False), encoding='utf-8')
    return path


def test_old_overrides_block_silent_success_and_upgrade_runs_every_step(tmp_path):
    event = old_event(tmp_path)
    store = lp_v3.Store(tmp_path)
    before = lp_v3.run_event(store, event, create_drafts=False, outdir=tmp_path/'out')
    assert before['status'] == 'preflight_failed'
    path = write_raw(store, event)
    original = path.read_bytes()
    ok, error, report = store.upgrade_event(event['id'], hashlib.sha256(original).hexdigest())
    assert ok, error
    assert Path(report['backup_path']).read_bytes() == original
    saved, error = store.load_event(event['id'])
    assert not error and not saved['process_overrides'] and not saved['process_results']
    assert saved['recipients'] == event['recipients']
    assert saved['pipeline'] == event['pipeline']
    run = lp_v3.run_event(store, saved, create_drafts=False, outdir=tmp_path/'out')
    assert run['status'] == 'ok' and all(len(row['files']) == 4 for row in run['items'])
    assert run['pipeline_stats']['executed'] == run['pipeline_stats']['enabled'] == 4
    receipt = next(store.runs_dir.glob('*' + run['id'] + '.json'))
    assert '流水线执行 4/4' in json.loads(receipt.read_text(encoding='utf-8'))['logs'][-1]


def test_hidden_attachments_upgrade_even_when_old_ui_marker_is_present(tmp_path):
    event, _ = make_event(tmp_path)
    event['pipeline_ui_version'] = '3.9'
    for index, recipient in enumerate(event['recipients']):
        recipient['attachments'] = [pdf(tmp_path/('own%d.pdf' % index))]
    event['unknown_metadata'] = {'preserve': 'original raw metadata'}
    store = lp_v3.Store(tmp_path)
    path = write_raw(store, event)
    ok, error, report = store.upgrade_event(event['id'])
    assert ok, error
    saved = json.loads(path.read_text(encoding='utf-8'))
    assert len(saved['pipeline']) == 2
    assert [{k:v for k,v in r.items() if k!='attachments'} for r in saved['recipients']] == [{k:v for k,v in r.items() if k!='attachments'} for r in event['recipients']]
    assert saved['unknown_metadata'] == event['unknown_metadata']
    assert saved['pipeline'][-1]['params']['path_source'] == 'item.paths'
    assert saved['pipeline'][-1]['params']['path_map'] == {r['name']:r['attachments'] for r in event['recipients']}
    assert all(not r['attachments'] for r in saved['recipients'])
    assert saved['pipeline'][-1]['params']['preserve_name']
    again, _ = lp_v3.upgrade_compatible_event(saved)
    assert len(again['pipeline']) == 2


def test_intentional_external_custom_files_are_preserved(tmp_path):
    event, _ = make_event(tmp_path)
    selected = pdf(tmp_path/'selected.pdf')
    event['process_overrides'] = {'LP_A': [selected]}
    event['process_results'] = {'LP_A': [selected]}
    upgraded, plan = lp_v3.upgrade_compatible_event(event)
    assert not plan['stale_override_names']
    assert upgraded['process_overrides'] == event['process_overrides']


@pytest.mark.parametrize('mutation', ['future_version', 'unknown_step', 'bad_structure'])
def test_unsafe_config_is_never_normalized_or_overwritten(tmp_path, mutation):
    event, _ = make_event(tmp_path)
    if mutation == 'future_version': event['config_version'] = '99.0.0'
    elif mutation == 'unknown_step': event['pipeline'][0]['kind'] = 'future-operation'
    else: event['pipeline'] = {'unknown': 'shape'}
    store = lp_v3.Store(tmp_path)
    path = write_raw(store, event)
    original = path.read_bytes()
    loaded, error = store.load_event(event['id'])
    assert loaded is None and error
    ok, error, report = store.upgrade_event(event['id'])
    assert not ok and error and report is None and path.read_bytes() == original
    safe, _ = make_event(tmp_path)
    safe['id'] = event['id']
    assert not store.save_event(safe)[0] and path.read_bytes() == original


@pytest.mark.parametrize('failure', ['backup', 'concurrent_change'])
def test_failed_upgrade_preserves_raw_event(tmp_path, monkeypatch, failure):
    store = lp_v3.Store(tmp_path)
    event = old_event(tmp_path)
    path = write_raw(store, event)
    original = path.read_bytes()
    if failure == 'backup':
        monkeypatch.setattr(store, '_backup_event', lambda *args: (_ for _ in ()).throw(OSError('backup failure')))
        ok, error, report = store.upgrade_event(event['id'])
    else:
        ok, error, report = store.upgrade_event(event['id'], 'incorrect hash')
    assert not ok and error and path.read_bytes() == original


def test_first_save_in_new_version_is_backed_up_once(tmp_path):
    event, _ = make_event(tmp_path)
    event['config_version'] = '3.9.1'
    event['pipeline_ui_version'] = '3.9'
    store = lp_v3.Store(tmp_path)
    path = write_raw(store, event)
    original = path.read_bytes()
    assert store.save_event(event)[0]
    backups = list((store.data_dir/'backups').glob('*.json'))
    assert len(backups) == 1 and backups[0].read_bytes() == original
    assert store.save_event(event)[0]
    assert len(list((store.data_dir/'backups').glob('*.json'))) == 1
    assert store.load_event(event['id'])[0]['config_version'] == lp_v3.APP_VERSION


@pytest.mark.parametrize('approve', [False, True])
def test_gui_assists_old_config_upgrade_and_cancel_preserves_file(app, tmp_path, monkeypatch, approve):
    event = old_event(tmp_path)
    path = write_raw(app.store, event)
    original = path.read_bytes()
    monkeypatch.setattr(lp_tool.messagebox, 'askyesno', lambda *args, **kwargs: approve)
    app.open_event(event['id'])
    if approve:
        assert not app._wb_event['process_overrides'] and app._wb_event['config_version'] == lp_v3.APP_VERSION
        assert list((app.store.data_dir/'backups').glob('*.json'))
    else:
        assert path.read_bytes() == original and app.cur_event_id != event['id']


def test_four_buttons_really_save_four_items(app, tmp_path):
    event, _ = make_event(tmp_path)
    event['pipeline'] = []
    show(app, event)
    for label in ('＋ Word MailMerge', '＋ PDF 文件', '＋ 非 PDF 文件', '＋ Excel 按 LP 匹配'):
        texts(app._wb_content)[label].invoke()
    stored = app.store.load_event(event['id'])[0]
    assert len(stored['pipeline']) == 4
    assert [lp_v3.item_input_type(item) for item in stored['pipeline']] == ['generate', 'pdf', 'file', 'excel']


def package(path, version='9.9.9', extra=None):
    with zipfile.ZipFile(path, 'w') as archive:
        for name in ('lp_tool.py', 'lp_v3.py'):
            archive.writestr('app/'+name, 'APP_VERSION = "%s"\n' % version)
        archive.writestr('app/watermark_encrypt.py', '# payload\n')
        archive.writestr('app/start.bat', '@echo off\n')
        for name, data in (extra or {}).items(): archive.writestr(name, data)
    return path


def collect_update(app):
    app._update_thread.join(timeout=10)
    assert not app._update_thread.is_alive()
    items = []
    while True:
        try: items.append(app._q.get_nowait())
        except queue.Empty: return items


def fake_download(monkeypatch, payload, total=None):
    def retrieve(url, filename, reporthook=None, **kwargs):
        reporthook(1, 1, total if total is not None else payload.stat().st_size)
        shutil.copyfile(payload, filename)
        reporthook(payload.stat().st_size, 1, total if total is not None else payload.stat().st_size)
    monkeypatch.setattr(urllib.request, 'urlretrieve', retrieve)


def test_update_reports_download_verify_extract_ready_and_preserves_data(app, tmp_path, monkeypatch):
    base = tmp_path/'installed'
    base.mkdir()
    (base/'lp_tool.py').write_bytes(b'old program')
    (base/'lp_data').mkdir()
    original = base/'lp_data'/'settings.json'
    original.write_bytes(b'original data')
    payload = package(tmp_path/'update.zip')
    fake_download(monkeypatch, payload)
    monkeypatch.setattr(lp_tool, 'app_dir', lambda: base)
    app._self_update({'tag': 'v9.9.9', 'asset_url': 'fake', 'asset_size': payload.stat().st_size,
                      'asset_digest': 'sha256:'+hashlib.sha256(payload.read_bytes()).hexdigest()})
    events = collect_update(app)
    assert any(kind == 'update_ready' for kind, _ in events)
    assert not any(kind == 'update_err' for kind, _ in events)
    phases = {data[0] for kind, data in events if kind == 'update_progress'}
    assert {'正在下载更新', '正在校验下载文件', '正在检查更新包', '正在解压更新', '正在准备安装', '更新文件已准备完成'} <= phases
    assert (base/'lp_tool.py').read_bytes() == b'old program' and original.read_bytes() == b'original data'
    assert (base/'run_update.ps1').exists()


@pytest.mark.parametrize('failure', ['digest', 'version', 'zip_path', 'user_data'])
def test_invalid_update_is_blocked_before_program_replacement(app, tmp_path, monkeypatch, failure):
    base = tmp_path/'installed';base.mkdir()
    current = base/'lp_tool.py';current.write_bytes(b'old')
    extra = {'../escaped.txt': 'bad'} if failure == 'zip_path' else {'app/lp_data/settings.json': 'bad'} if failure == 'user_data' else {}
    payload = package(tmp_path/'bad.zip', '9.9.8' if failure == 'version' else '9.9.9', extra)
    fake_download(monkeypatch, payload)
    monkeypatch.setattr(lp_tool, 'app_dir', lambda: base)
    app._self_update({'tag': 'v9.9.9', 'asset_url': 'fake', 'asset_digest': 'sha256:wrong' if failure == 'digest' else ''})
    events = collect_update(app)
    assert any(kind == 'update_err' for kind, _ in events)
    assert not any(kind == 'update_ready' for kind, _ in events)
    assert current.read_bytes() == b'old' and not (tmp_path/'escaped.txt').exists()


def test_unknown_download_total_is_indeterminate(app, tmp_path, monkeypatch):
    base=tmp_path/'installed';base.mkdir()
    payload=package(tmp_path/'update.zip')
    fake_download(monkeypatch, payload, -1)
    monkeypatch.setattr(lp_tool, 'app_dir', lambda: base)
    app._self_update({'tag': 'v9.9.9', 'asset_url': 'fake'})
    events=collect_update(app)
    assert any(kind=='update_progress' and data[0]=='正在下载更新' and data[1] is None for kind,data in events)


def test_progress_window_visible_cancel_works_and_duplicate_start_is_blocked(app, tmp_path, monkeypatch):
    base=tmp_path/'installed';base.mkdir()
    started=threading.Event();release=threading.Event()
    def retrieve(url, filename, reporthook=None, **kwargs):
        reporthook(1, 1024, 4096);started.set();release.wait(5);reporthook(2,1024,4096)
    monkeypatch.setattr(lp_tool, 'app_dir', lambda: base)
    monkeypatch.setattr(urllib.request, 'urlretrieve', retrieve)
    app.check_update=lambda *args, **kwargs: None
    app.root.deiconify()
    assert app._self_update({'tag':'v9.9.9','asset_url':'fake'})
    assert started.wait(2)
    deadline=time.monotonic()+3
    while time.monotonic()<deadline:
        app.root.update()
        if '已下载' in app._update_detail_var.get(): break
        time.sleep(.02)
    assert app._update_window.winfo_ismapped()
    assert '已下载' in app._update_detail_var.get()
    assert not app._self_update({'tag':'v9.9.9','asset_url':'fake'})
    app._update_cancel_button.invoke();release.set()
    events=collect_update(app)
    assert any(kind=='update_cancelled' for kind,_ in events)
    assert not (base/'.update_pkg.zip').exists()


@pytest.mark.parametrize('fail_install', [False, True])
def test_native_installer_backup_copy_and_failure_restore(tmp_path, fail_install):
    base=tmp_path/'app';base.mkdir()
    source=base/'.update_staging'/'app';source.mkdir(parents=True)
    current=base/'a.py';current.write_bytes(b'old program')
    (source/'a.py').write_bytes(b'new program')
    data=base/'lp_data'/'events';data.mkdir(parents=True)
    (data/'event.json').write_bytes(b'private original')
    history=base/'lp_data'/'backups';history.mkdir();(history/'historic.json').write_bytes(b'old backup')
    old_update=base/'.update_backups'/'old';old_update.mkdir(parents=True);(old_update/'keep.txt').write_bytes(b'previous backup')
    (base/'lp_data.json.bak').write_bytes(b'legacy backup')
    seed=base/'lp_data.example.json';seed.write_bytes(b'user seed')
    (source/'lp_data.example.json').write_bytes(b'new seed must not overwrite')
    if fail_install:
        (base/'z').write_bytes(b'file blocks directory')
        (source/'z').mkdir();(source/'z'/'child.py').write_bytes(b'child')
    helper=lp_tool._write_update_helper(base, source)
    run=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(helper),
                        '-WaitProcessId','0','-NoUi','-NoRestart'],capture_output=True,text=True,timeout=30)
    result=json.loads((base/'.update_result.json').read_text(encoding='utf-8-sig'))
    assert result['success'] is not fail_install, (run.stdout,run.stderr)
    assert (data/'event.json').read_bytes()==b'private original' and seed.read_bytes()==b'user seed'
    assert current.read_bytes()==(b'old program' if fail_install else b'new program')
    program_backup=Path(result['backup'])/'program'/'a.py'
    if fail_install:
        assert program_backup.read_bytes()==b'old program'
    else:
        assert not program_backup.exists() and result['program_backup_removed']
    assert (Path(result['backup'])/'lp_data'/'events'/'event.json').read_bytes()==b'private original'
    assert (Path(result['backup'])/'legacy'/'lp_data.json.bak').read_bytes()==b'legacy backup'
    assert (history/'historic.json').read_bytes()==b'old backup'
    assert not (Path(result['backup'])/'lp_data'/'backups').exists()
    assert (old_update/'keep.txt').read_bytes()==b'previous backup'


def test_installer_hash_mismatch_rolls_back_and_retains_backup(tmp_path):
    base=tmp_path/'app';base.mkdir()
    source=base/'.update_staging'/'app';source.mkdir(parents=True)
    (base/'a.py').write_bytes(b'old program');(source/'a.py').write_bytes(b'new program')
    data=base/'lp_data'/'events';data.mkdir(parents=True);(data/'event.json').write_bytes(b'unchanged')
    helper=lp_tool._write_update_helper(base,source)
    script=helper.read_text(encoding='utf-8-sig')
    needle="    $index=0\n    foreach ($file in $files) {\n        $relative=$file.FullName.Substring($source.Length+1);$target=Destination $relative\n        if ((FileHash"
    assert needle in script
    helper.write_text(script.replace(needle,"    Set-Content -LiteralPath (Join-Path $root 'a.py') -Value 'tampered'\n"+needle,1),encoding='utf-8-sig')
    run=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(helper),
                        '-WaitProcessId','0','-NoUi','-NoRestart'],capture_output=True,text=True,timeout=30)
    receipt=json.loads((base/'.update_result.json').read_text(encoding='utf-8-sig'))
    assert run.returncode != 0 and not receipt['success']
    assert (base/'a.py').read_bytes()==b'old program'
    assert (Path(receipt['backup'])/'program'/'a.py').read_bytes()==b'old program'
    assert (data/'event.json').read_bytes()==b'unchanged'


def test_installer_dependency_check_failure_rolls_back(tmp_path):
    base=tmp_path/'app';base.mkdir()
    source=base/'.update_staging'/'app';source.mkdir(parents=True)
    (base/'lp_tool.py').write_bytes(b'old program')
    (source/'lp_tool.py').write_bytes(b'new program')
    data=base/'lp_data'/'events';data.mkdir(parents=True);(data/'event.json').write_bytes(b'unchanged')
    helper=lp_tool._write_update_helper(base,source)
    run=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(helper),
                        '-WaitProcessId','0','-NoUi','-NoRestart'],capture_output=True,text=True,timeout=30)
    receipt=json.loads((base/'.update_result.json').read_text(encoding='utf-8-sig'))
    assert run.returncode != 0 and not receipt['success']
    assert (base/'lp_tool.py').read_bytes()==b'old program'
    assert (Path(receipt['backup'])/'program'/'lp_tool.py').read_bytes()==b'old program'
    assert (data/'event.json').read_bytes()==b'unchanged'


@pytest.mark.parametrize('extension', ['.xlsx', '.pptx'])
def test_office_pdf_conversion_owns_separate_com_thread(tmp_path, monkeypatch, extension):
    import watermark_encrypt as wm
    pythoncom = pytest.importorskip('pythoncom')
    client = pytest.importorskip('win32com.client')
    calls = []
    def record(name): calls.append((name, threading.get_ident()))
    class Document:
        def ExportAsFixedFormat(self, *args): record('export')
        def SaveAs(self, *args): record('export')
        def Close(self, *args): record('close')
    class Documents:
        def Open(self, *args, **kwargs): record('open');return Document()
    class Office:
        def __init__(self): self.Workbooks = self.Presentations = Documents()
        def Quit(self): record('quit')
    monkeypatch.setattr(pythoncom, 'CoInitializeEx', lambda *args: record('initialize'))
    monkeypatch.setattr(pythoncom, 'CoUninitialize', lambda: record('uninitialize'))
    monkeypatch.setattr(client, 'DispatchEx', lambda *args: Office())
    caller = threading.get_ident()
    wm.file_to_pdf(str(tmp_path/('input'+extension)), str(tmp_path/'out.pdf'))
    assert [name for name, _ in calls] == ['initialize', 'open', 'export', 'close', 'quit', 'uninitialize']
    assert all(thread != caller for _, thread in calls)
    assert len({thread for _, thread in calls}) == 1


@pytest.mark.parametrize('action', ['disable', 'delete'])
def test_migrated_attachment_cannot_reappear_when_item_disabled_or_deleted(tmp_path, action):
    event, _ = make_event(tmp_path)
    event['pipeline_ui_version'] = '3.9'
    for index, recipient in enumerate(event['recipients']):
        recipient['attachments'] = [pdf(tmp_path/('legacy%d.pdf'%index), 'OLD%d'%index)]
    upgraded, _ = lp_v3.upgrade_compatible_event(event)
    assert all(not recipient['attachments'] for recipient in upgraded['recipients'])
    assert len(upgraded['pipeline']) == 2
    if action == 'disable': upgraded['pipeline'][-1]['enabled'] = False
    else: upgraded['pipeline'].pop()
    again, plan = lp_v3.upgrade_compatible_event(upgraded)
    assert len(again['pipeline']) == len(upgraded['pipeline']) and not plan.get('needs_confirmation')
    run=lp_v3.run_event(lp_v3.Store(tmp_path),again,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok' and all(len(row['files'])==1 for row in run['items'])


def test_disabled_old_reference_is_converted_without_reenabling_it(tmp_path):
    event,_=make_event(tmp_path)
    item=lp_v3.new_item('excel');item['enabled']=False
    item['params']['path_source']=lp_v3.PATH_SOURCE_RECIPIENT;event['pipeline'].append(item)
    for recipient in event['recipients']:recipient['attachments']='a.pdf;b.pdf'
    upgraded, _=lp_v3.upgrade_compatible_event(event)
    assert len(upgraded['pipeline'])==2 and not upgraded['pipeline'][-1]['enabled']
    assert upgraded['pipeline'][-1]['params']['path_map']['LP_A']==['a.pdf','b.pdf']
    assert not upgraded['recipients'][0]['attachments']


def test_ambiguous_old_and_new_excel_relationships_require_manual_review(tmp_path):
    event,item=make_event(tmp_path,'excel')
    item['params']['path_map']={'LP_A':['new.pdf']}
    event['recipients'][0]['attachments']=['old.pdf']
    store=lp_v3.Store(tmp_path);path=write_raw(store,event);before=path.read_bytes()
    ok,error,report=store.upgrade_event(event['id'])
    assert not ok and '无法自动判断' in error and path.read_bytes()==before
