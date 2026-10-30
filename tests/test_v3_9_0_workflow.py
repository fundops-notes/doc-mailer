"""Four workflow entries, final file selections, Next actions and compact editors."""
from pathlib import Path
import json
import sys
import fitz
import openpyxl
import pytest
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import lp_tool
import lp_v3
import watermark_encrypt as wm
from test_v3_pool_gui import app, no_messagebox, _walk


def pdf(path, text='TEST'):
    document=fitz.open();page=document.new_page();page.insert_text((40,50),text)
    document.save(path);document.close();return str(path)


def make_event(tmp_path, kind='pdf'):
    event=lp_v3.new_event('工作流验收')
    event['recipients']=[lp_v3.normalize_recipient({'name':name,'to':name+'@example.com',
        'open_pw':'open-'+name,'owner_pw':'owner-'+name}) for name in ('LP_A','LP_B')]
    item=lp_v3.new_item(kind)
    if kind!='generate': item['params']['files']=[pdf(tmp_path/'input.pdf')]
    event['pipeline']=[item]
    return event,item


def show(a,event,step=1):
    a.check_update=lambda *args,**kw:None
    a.store.save_event(event);a.open_event(event['id']);a._wb_show_step(step)
    a.root.deiconify();a.root.geometry('1280x800')
    for _ in range(4):a.root.update()
    return a._wb_event


def texts(host):
    return {str(w.cget('text')):w for w in _walk(host) if w.winfo_class() in ('TButton','TMenubutton')}


@pytest.mark.parametrize('variant',['generate','pdf','file','excel'])
def test_four_variants_normalize_roundtrip(variant):
    item=lp_v3.new_item(variant)
    assert lp_v3.item_input_type(lp_v3.normalize_item(item))==variant
    assert item['kind']==('generate' if variant=='generate' else 'attach')


def test_four_entry_buttons_and_no_separate_process_or_common_attachment(app,tmp_path):
    event,_=make_event(tmp_path);show(app,event)
    buttons=texts(app._wb_content)
    assert {'＋ Word MailMerge','＋ PDF 文件','＋ 非 PDF 文件','＋ Excel 按 LP 匹配'}<=buttons.keys()
    assert not any('统一附件' in text or text.startswith('▶') for text in buttons)
    assert '处理' in app._wb_next_btn.cget('text')


@pytest.mark.parametrize('variant',['pdf','file','excel'])
def test_parameters_match_selected_entry(app,tmp_path,variant):
    event,item=make_event(tmp_path,variant);show(app,event)
    labels=[str(w.cget('text')) for w in _walk(app._wb_params_frame) if w.winfo_class()=='TCheckbutton']
    assert ('转成 PDF' in labels)==(variant!='pdf')
    buttons=texts(app._wb_params_frame)
    assert ('导入对应清单…' in buttons)==(variant=='excel')
    assert not ('从 Excel 导入路径…' in buttons)


def test_legacy_attachments_surface_in_pipeline_and_backup_before_save(app,tmp_path):
    event,item=make_event(tmp_path)
    common=pdf(tmp_path/'common.pdf')
    event['attachments']=[lp_v3.normalize_attachment({'path':common,'scope':'all'})]
    app.store.save_event(event)
    original=app.store.event_path(event['id']).read_bytes()
    app.open_event(event['id'])
    assert app._wb_event['attachments']==[]
    assert any(i['params'].get('files')==[common] for i in app._wb_event['pipeline'])
    assert app.store.event_path(event['id']).read_bytes()==original
    assert app._wb_save()
    backups=list((app.store.data_dir/'backups').glob('*.pre_v3_9_0_*.json'))
    assert len(backups)==1 and backups[0].read_bytes()==original
    assert app._wb_save()
    assert len(list((app.store.data_dir/'backups').glob('*.json')))==1


def test_legacy_backup_failure_does_not_overwrite_event(app,tmp_path,monkeypatch):
    import shutil
    event,_=make_event(tmp_path);app.store.save_event(event)
    original=app.store.event_path(event['id']).read_bytes()
    migrated=lp_v3.upgrade_pipeline_event(event)
    monkeypatch.setattr(shutil,'copy2',lambda *a,**k:(_ for _ in ()).throw(OSError('backup failed')))
    ok,error=app.store.save_event(migrated)
    assert not ok and '备份' in error
    assert app.store.event_path(event['id']).read_bytes()==original


def test_excel_maps_are_item_local_and_support_relative_paths(app,tmp_path,monkeypatch):
    event,_=make_event(tmp_path)
    first=lp_v3.new_item('excel');second=lp_v3.new_item('excel')
    second['params']['path_map']={'LP_A':['other.pdf']}
    event['pipeline']=[first,second];show(app,event)
    book=openpyxl.Workbook();sheet=book.active
    sheet.append(['收件人名称','附件路径']);sheet.append(['LP_A','a.pdf']);sheet.append(['LP_A','b.pdf']);sheet.append(['LP_B','c.pdf'])
    path=tmp_path/'mapping.xlsx';book.save(path);book.close()
    monkeypatch.setattr(lp_tool.filedialog,'askopenfilename',lambda **kw:str(path))
    item=app._wb_event['pipeline'][0]
    app._wb_import_attach_paths(item,None)
    dialog=next(w for w in app.root.winfo_children() if w.winfo_class()=='Toplevel')
    texts(dialog)['导入'].invoke()
    assert item['params']['path_map']['LP_A']==[str((tmp_path/'a.pdf').resolve()),str((tmp_path/'b.pdf').resolve())]
    assert app._wb_event['pipeline'][1]['params']['path_map']=={'LP_A':['other.pdf']}
    assert all(not r['attachments'] for r in app._wb_event['recipients'])
    stored=app.store.load_event(event['id'])[0]
    assert stored['pipeline'][0]['params']['path_map']==item['params']['path_map']


def test_nonpdf_unchanged_when_conversion_unchecked(tmp_path):
    event,item=make_event(tmp_path,'file')
    raw=tmp_path/'input.zip';raw.write_bytes(b'raw attachment bytes')
    item['params']['files']=[str(raw)]
    run=lp_v3.run_event(lp_v3.Store(tmp_path),event,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok'
    assert all(Path(i['files'][0]).suffix=='.zip' and Path(i['files'][0]).read_bytes()==raw.read_bytes() for i in run['items'])


def test_nonpdf_conversion_then_watermark_encryption(tmp_path):
    event,item=make_event(tmp_path,'file')
    image=tmp_path/'photo.png';Image.new('RGB',(160,120),'white').save(image)
    item['params'].update(files=[str(image)],to_pdf=True,watermark=True,encrypt=True)
    run=lp_v3.run_event(lp_v3.Store(tmp_path),event,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok',run['errors']
    for recipient,result in zip(event['recipients'],run['items']):
        with fitz.open(result['files'][0]) as document:
            assert document.needs_pass
            assert document.authenticate(recipient['open_pw'])
            assert recipient['name'] in document[0].get_text()


def test_colliding_converted_names_do_not_overwrite(tmp_path):
    event,item=make_event(tmp_path,'file')
    first=tmp_path/'report.png';second=tmp_path/'report.jpg'
    Image.new('RGB',(100,80),'red').save(first);Image.new('RGB',(140,90),'green').save(second)
    item['params'].update(files=[str(first),str(second)],to_pdf=True)
    run=lp_v3.run_event(lp_v3.Store(tmp_path),event,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok'
    assert all(len(i['files'])==2 and len(set(i['files']))==2 for i in run['items'])


@pytest.mark.parametrize('variant,to_pdf,watermark', [('pdf',False,False),('file',True,False),('file',False,True)])
def test_invalid_conversion_combinations_block_preflight(tmp_path,variant,to_pdf,watermark):
    event,item=make_event(tmp_path,variant)
    source=tmp_path/'unsupported.zip';source.write_bytes(b'test')
    item['params'].update(files=[str(source)],to_pdf=to_pdf,watermark=watermark)
    errors,_,_=lp_v3.preflight(event,lp_v3.Store(tmp_path))
    assert errors


def test_saved_overrides_override_final_files_even_without_valid_template(tmp_path):
    event,item=make_event(tmp_path,'generate')
    event['process_overrides']={r['name']:[pdf(tmp_path/(r['name']+'.pdf'),r['name'])] for r in event['recipients']}
    run=lp_v3.run_event(lp_v3.Store(tmp_path),event,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok',run['errors']
    assert all(i['files']==event['process_overrides'][i['name']] for i in run['items'])


def test_prepared_files_reused_for_drafts_and_invalidated_on_source_change(tmp_path,monkeypatch):
    event,item=make_event(tmp_path)
    store=lp_v3.Store(tmp_path)
    processed=lp_v3.run_event(store,event,create_drafts=False,outdir=tmp_path/'out')
    event['process_results']={i['name']:i['files'] for i in processed['items']}
    event['process_signature']=processed['process_signature']
    assert lp_v3.prepared_files_current(event)
    calls=[]
    monkeypatch.setattr(lp_v3,'create_outlook_draft',lambda recipient,files,*args,**kwargs:calls.append((recipient['name'],[str(p) for p in files])))
    monkeypatch.setattr(lp_v3,'run_items',lambda *a,**k:(_ for _ in ()).throw(AssertionError('must not rerun')))
    run=lp_v3.run_event(store,event,reuse_processed=True,outdir=tmp_path/'out')
    assert run['status']=='ok' and len(calls)==2
    assert all(files==event['process_results'][name] for name,files in calls)
    source=Path(item['params']['files'][0]);source.write_bytes(source.read_bytes()+b'\nchanged')
    assert not lp_v3.prepared_files_current(event)


def test_next_commits_latest_inputs_and_waits_for_processing_success(app,tmp_path,monkeypatch):
    event,item=make_event(tmp_path);show(app,event)
    replacement=pdf(tmp_path/'latest.pdf')
    editor=app._wb_param_editor
    editor['files'].delete('1.0','end');editor['files'].insert('1.0',replacement)
    editor['name'].set('最新名称')
    monkeypatch.setattr(app,'_default_outdir',lambda:tmp_path/'output')
    monkeypatch.setattr(app,'_submit_job',lambda _job:None)
    app._wb_next_btn.invoke()
    assert app.running and app._wb_step==1
    assert app._wb_event['pipeline'][0]['params']['files']==[replacement]
    assert app.store.load_event(event['id'])[0]['pipeline'][0]['name']=='最新名称'
    run={'status':'ok','stats':{'ok':2,'fail':0,'skip':0},'errors':[],
         'items':[{'name':r['name'],'members':[r['name']],'files':[replacement]} for r in event['recipients']],
         'process_signature':lp_v3.processing_signature(app._wb_event)}
    app._wb_finish_run((run,lp_v3.normalize_event(app._wb_event)))
    assert app._wb_step==2 and not app.running


def test_processing_failure_stays_on_pipeline(app,tmp_path,monkeypatch):
    event,_=make_event(tmp_path);show(app,event)
    monkeypatch.setattr(app,'_default_outdir',lambda:tmp_path/'output')
    monkeypatch.setattr(app,'_submit_job',lambda _job:None)
    app._wb_next_btn.invoke()
    app._wb_finish_run(({'status':'partial','errors':['conversion failed'],'stats':{'fail':1},'items':[]},app._wb_event))
    assert app._wb_step==1 and not app.running


@pytest.mark.parametrize('confirm',[False,True])
def test_mail_next_prompts_save_then_previews(app,tmp_path,monkeypatch,confirm):
    event,_=make_event(tmp_path);event['mail_overrides']={'old':{'subject':'old','body':'old'}}
    show(app,event,2);app.mail_subject_var.set('尚未保存的新主题')
    prompts=[]
    monkeypatch.setattr(lp_tool.messagebox,'askokcancel',lambda *args,**kw:prompts.append(args) or confirm)
    app._wb_next_btn.invoke()
    assert prompts
    assert app._wb_step==(3 if confirm else 2)
    if confirm:
        assert app.store.load_event(event['id'])[0]['mail']['subject']=='尚未保存的新主题'
        assert app._wb_event['mail_overrides']=={}
    else:
        assert app._mail_dirty and app._wb_event['mail_overrides']
    assert '预生成 →' not in texts(app._wb_content)


def test_failed_save_does_not_preview_or_clear_single_mail_edits(app,tmp_path,monkeypatch):
    event,_=make_event(tmp_path);event['mail_overrides']={'old':{'subject':'old','body':'old'}}
    show(app,event,2);app.mail_subject_var.set('new')
    monkeypatch.setattr(app.store,'save_event',lambda *args:(False,'disk unavailable'))
    app._wb_next_btn.invoke()
    assert app._wb_step==2 and app._mail_dirty and app._wb_event['mail_overrides']


@pytest.mark.parametrize('step',[2,3])
def test_word_style_toolbar_is_compact_in_both_editors(app,tmp_path,step):
    from tkinter import ttk
    event,_=make_event(tmp_path);show(app,event,step)
    controls=texts(app._editor_toolbar)
    assert {'B','I','U','文字色','背景色','左对齐','居中','右对齐','插入图片…','清除格式','插入字段 ▾'}<=controls.keys()
    reference=ttk.Button(app._wb_content,text='默认按钮')
    assert controls['文字色'].winfo_reqheight()<reference.winfo_reqheight()
    assert all(controls[name].cget('image') for name in ('左对齐','居中','右对齐'))
    reference.destroy()


@pytest.mark.parametrize('column',['#2','#3'])
def test_mapping_file_and_path_columns_edit_and_next_commits(app,tmp_path,monkeypatch,column):
    from types import SimpleNamespace
    event,_=make_event(tmp_path);show(app,event)
    tree=app._wb_map_tree
    row=tree.get_children()[0]
    tree.xview_moveto(0 if column=='#2' else 1)
    app.root.update()
    x,y,w,h=tree.bbox(row,column)
    app._wb_map_edit(SimpleNamespace(x=x+max(5,w//2),y=y+h//2))
    editor=app._wb_map_editor
    selected=pdf(tmp_path/'custom.pdf')
    editor.delete(0,'end');editor.insert(0,selected)
    monkeypatch.setattr(app,'_default_outdir',lambda:tmp_path/'out')
    monkeypatch.setattr(app,'_submit_job',lambda _job:None)
    app._wb_next_btn.invoke()
    assert app.running and app._wb_step==1
    assert app.store.load_event(event['id'])[0]['process_overrides']['LP_A']==[selected]


def test_partial_retry_preserves_other_recipient_prepared_files(app,tmp_path):
    event,_=make_event(tmp_path)
    paths={name:[pdf(tmp_path/(name+'.pdf'),name)] for name in ('LP_A','LP_B')}
    event['process_results']=paths
    show(app,event)
    run={'status':'ok','only_recips':['LP_A'],'errors':[],
         'process_signature':lp_v3.processing_signature(app._wb_event),
         'items':[{'name':'LP_A','members':['LP_A'],'files':paths['LP_A']}]}
    assert app._apply_process_results(run)
    assert app._wb_event['process_results']==paths
    assert lp_v3.prepared_files_current(app._wb_event)


def test_incomplete_prepared_files_never_count_as_current(tmp_path):
    event,_=make_event(tmp_path)
    event['process_results']={'LP_A':[pdf(tmp_path/'a.pdf')]}
    event['process_signature']=lp_v3.processing_signature(event)
    assert not lp_v3.prepared_files_current(event)


def test_legacy_field_path_file_change_invalidates_processing_signature(tmp_path):
    event,item=make_event(tmp_path,'excel')
    source=pdf(tmp_path/'from_field.pdf')
    item['params']['path_source']='fields.File'
    for recipient in event['recipients']: recipient['fields']['File']=source
    signature=lp_v3.processing_signature(event)
    with open(source,'ab') as output: output.write(b'\nmodified')
    assert lp_v3.processing_signature(event)!=signature
