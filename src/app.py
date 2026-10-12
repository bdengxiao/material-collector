import sys, json, copy
from pathlib import Path
from PySide6.QtCore import QThread,Signal,Qt
from PySide6.QtGui import QFontDatabase,QFont
from PySide6.QtWidgets import *
from core import *

class Worker(QThread):
    done=Signal(object); failed=Signal(str); progress=Signal(int)
    def __init__(self,fn): super().__init__(); self.fn=fn
    def run(self):
        try: self.done.emit(self.fn(self.isInterruptionRequested,self.progress.emit))
        except Exception as e: self.failed.emit(str(e))

class Window(QMainWindow):
    def __init__(self):
        super().__init__(); self.project=Project(); self.project_path=None; self.worker=None; self.saved_snapshot=None
        for name in ('msyh.ttc','msyhbd.ttc'):
            font=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'/name
            if font.exists():QFontDatabase.addApplicationFont(str(font))
        self.setFont(QFont('Microsoft YaHei',10))
        self.setWindowTitle('材料收齐助手 0.2.2 · 试用候选'); self.resize(1180,760); self.setMinimumSize(860,600)
        self.setStyleSheet('QWidget{font-family:Microsoft YaHei;font-size:14px;color:#243447} QMainWindow,QStackedWidget{background:#f4f6f9} QPushButton{padding:10px 16px;border:1px solid #cbd5e1;border-radius:5px;background:white} QPushButton:hover{background:#e7eef8} QPushButton:focus{border:2px solid #245eae} QLineEdit,QComboBox,QSpinBox{padding:7px;background:white;border:1px solid #b8c5d5} QTableWidget{background:white;gridline-color:#e5eaf0} QHeaderView::section{background:#e7edf5;padding:9px;border:0;font-weight:bold} QListWidget{background:#183153;color:white;border:0} QListWidget::item{padding:17px} QListWidget::item:selected{background:#28558b}')
        central=QWidget(); self.setCentralWidget(central); layout=QHBoxLayout(central); layout.setContentsMargins(0,0,0,0)
        self.nav=QListWidget(); self.nav.setFixedWidth(180); self.nav.addItems(['首页','名单与规则','核对结果','人工确认','整理导出','历史与保存','帮助']); layout.addWidget(self.nav)
        self.pages=QStackedWidget(); layout.addWidget(self.pages); self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.home(); self.setup(); self.results(); self.manual_page(); self.export_page(); self.history(); self.help_page(); self.nav.setCurrentRow(0)
        self.bar=QProgressBar(); self.bar.setMaximumHeight(14); self.statusBar().addPermanentWidget(self.bar); self.statusBar().showMessage('所有材料在本机处理；原文件保持不变')
    def page(self,title,subtitle):
        w=QWidget(); l=QVBoxLayout(w); l.setContentsMargins(28,24,28,24); h=QLabel(title); h.setStyleSheet('font-size:25px;font-weight:600'); l.addWidget(h); s=QLabel(subtitle); s.setWordWrap(True); l.addWidget(s); self.pages.addWidget(w); return l
    def button(self,l,text,fn):
        b=QPushButton(text); b.clicked.connect(fn); l.addWidget(b); return b
    def notice(self,text): QMessageBox.information(self,'材料收齐助手',text)
    def guarded(self,fn):
        try: fn()
        except Exception as e: QMessageBox.warning(self,'操作未完成',str(e))
    def idle(self):
        if self.worker and self.worker.isRunning():raise ValueError('正在处理，请等待完成或取消后再修改项目。')
    def discard_confirm(self):
        return not self.project.people or QMessageBox.question(self,'切换任务','当前任务可能有未保存的修改。请先保存，或确认放弃未保存修改。继续切换？')==QMessageBox.Yes
    def sync_rules(self):
        self.idle()
        rules=self.get_rules()
        if self.project.rules!=rules:
            self.notice('规则变化只要求受影响的人工确认复核，原归属保留供参考；未受影响确认保持不变。')
            update_rules(self.project,rules)
    def home(self):
        l=self.page('把材料收齐，少一些挨个核对','导入名单和已收到的文件夹，查看缺件并人工处理疑问。文件已找到不代表内容审核合格。')
        self.button(l,'新建核对任务',self.new); self.button(l,'打开历史任务',self.open_project); self.button(l,'使用 10 人虚构示例',self.example); self.button(l,'三分钟体验：照着五步操作',self.quick_guide); self.button(l,'查看使用教程',lambda:self.nav.setCurrentRow(6)); l.addStretch()
    def quick_guide(self):
        self.notice('① 首页点“使用 10 人虚构示例”：10人、17文件、3人缺件。\n② 人工确认页选“匿名照片.png”，指定示例人员8（0008）和照片：缺件变2人。\n③ 历史与保存页点“保存项目”，保存为自己的“体验.msq”。\n④ 整理导出页选择材料目录之外的独立文件夹，先预览再导出：3份Excel和ZIP；缺件人员9、10。\n⑤ 关闭后重新启动，打开体验.msq，检查人工归属仍在。\n\n仅使用虚构示例。保存项目不备份原材料；请保留源文件位置。找到文件不表示内容合格。')
    def new(self):
        if self.worker and self.worker.isRunning():self.notice('请先等待或取消当前操作。');return
        if not self.discard_confirm():return
        self.project=Project(); self.project_path=None;self.saved_snapshot=None;self.apply_template();self.refresh(); self.nav.setCurrentRow(1)
    def setup(self):
        l=self.page('1 · 名单与材料规则','先导入名单，再配置每个人需要提交的材料。编号建议使用文本格式。')
        self.button(l,'导入 Excel / CSV 名单（预览并选择列）',lambda:self.guarded(self.import_roster))
        self.roster_label=QLabel('尚未导入名单'); l.addWidget(self.roster_label)
        self.template=QComboBox(); self.template.addItems(TEMPLATES); l.addWidget(self.template)
        self.button(l,'应用内置模板',lambda:self.guarded(self.apply_template))
        self.rule_table=QTableWidget(0,5); self.rule_table.setHorizontalHeaderLabels(['材料名称','文件扩展名（逗号分隔）','必交','可多份','关键词（逗号分隔）']); self.rule_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); l.addWidget(self.rule_table)
        row=QHBoxLayout(); l.addLayout(row)
        for text,fn in [('新增材料',self.add_rule),('删除选中规则',self.remove_rule),('保存规则模板',self.save_template),('载入规则模板',self.load_template)]:
            b=QPushButton(text); b.clicked.connect(lambda checked=False,f=fn:self.guarded(f)); row.addWidget(b)
        self.button(l,'确认规则并选择材料目录 → 开始核对',lambda:self.guarded(self.start_scan))
        self.apply_template()
    def import_roster(self):
        self.idle()
        p,_=QFileDialog.getOpenFileName(self,'导入名单','','名单 (*.xlsx *.csv)')
        if not p:return
        headers,rows,warnings=read_roster(p); dlg=QDialog(self); dlg.setWindowTitle('名单预览 · 只保存姓名与所选编号'); dlg.resize(760,500); l=QVBoxLayout(dlg)
        names=QComboBox(); names.addItems(headers); ids=QComboBox(); ids.addItems(['不使用编号']+headers)
        for i,h in enumerate(headers):
            if h in ('姓名','名字'): names.setCurrentIndex(i)
            if h in ('学号','工号','编号'): ids.setCurrentIndex(i+1)
        l.addWidget(QLabel('姓名列')); l.addWidget(names); l.addWidget(QLabel('编号列（手机号后四位不作为身份编号）'));l.addWidget(ids)
        table=QTableWidget(min(len(rows),20),len(headers)); table.setHorizontalHeaderLabels(headers)
        for i,r in enumerate(rows[:20]):
            for j,v in enumerate(r[:len(headers)]): table.setItem(i,j,QTableWidgetItem(str(v)))
        l.addWidget(table); l.addWidget(QLabel('\n'.join(warnings) or f'共有 {len(rows)} 行，仅预览前 20 行'))
        ok=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel); ok.accepted.connect(dlg.accept);ok.rejected.connect(dlg.reject);l.addWidget(ok)
        if dlg.exec()!=QDialog.Accepted:return
        people,extra=parse_people(headers,rows,names.currentIndex(),ids.currentIndex()-1 if ids.currentIndex() else None)
        if self.project.materials and QMessageBox.question(self,'更换名单','更换名单会清空已导入材料及人工确认。已保存项目仍可打开，继续？')!=QMessageBox.Yes:return
        self.project.people=people;self.project.materials=[];self.project.roots=[];self.project.event('导入名单',count=len(people));self.refresh()
        if warnings+extra:self.notice('\n'.join(warnings+extra))
    def add_rule(self,r=None):
        self.idle()
        if not isinstance(r,Rule):r=Rule('新材料',['.pdf'])
        i=self.rule_table.rowCount();self.rule_table.insertRow(i)
        for j,v in enumerate([r.name,','.join(r.extensions),'是' if r.required else '否','是' if r.multiple else '否',','.join(r.keywords)]):self.rule_table.setItem(i,j,QTableWidgetItem(v))
    def remove_rule(self):
        self.idle()
        if self.rule_table.currentRow()>=0:self.rule_table.removeRow(self.rule_table.currentRow())
    def apply_template(self):
        self.idle()
        self.rule_table.setRowCount(0)
        for r in TEMPLATES[self.template.currentText()]:self.add_rule(r)
    def get_rules(self):
        rules=[]
        for i in range(self.rule_table.rowCount()):
            values=[self.rule_table.item(i,j).text().strip() if self.rule_table.item(i,j) else '' for j in range(5)]
            name,ext,required,multiple,kw=values
            extensions=[x.strip().lower() for x in ext.replace('，',',').split(',') if x.strip()]
            if not name or not extensions or any(not re.fullmatch(r'\.[a-z0-9]+',x) for x in extensions):raise ValueError('每种材料需要名称及 .pdf 等扩展名。')
            if required not in ('是','否') or multiple not in ('是','否'):raise ValueError('必交和可多份请填写“是”或“否”。')
            rules.append(Rule(name,extensions,required=='是',multiple=='是',[x.strip() for x in kw.replace('，',',').split(',') if x.strip()]))
        if not rules or len({r.name for r in rules})!=len(rules):raise ValueError('至少一条规则，且材料名称不能重复。')
        return rules
    def save_template(self):
        self.idle()
        rules=self.get_rules();p,_=QFileDialog.getSaveFileName(self,'保存规则模板','','规则 (*.json)')
        if p:
            reject_link_path(p)
            with Path(p).open('x',encoding='utf-8') as out:out.write(json.dumps([asdict(r) for r in rules],ensure_ascii=False,indent=2))
    def load_template(self):
        self.idle()
        p,_=QFileDialog.getOpenFileName(self,'载入规则模板','','规则 (*.json)')
        if p:
            rules=[Rule(**x) for x in json.loads(Path(p).read_text(encoding='utf-8-sig'))];self.rule_table.setRowCount(0)
            for r in rules:self.add_rule(r)
    def run_job(self,fn,done):
        if self.worker and self.worker.isRunning():raise ValueError('正在处理，请等待或取消。')
        self.rule_table.setEnabled(False);self.template.setEnabled(False)
        self.worker=Worker(fn); self.bar.setRange(0,0);self.worker.progress.connect(lambda n:self.statusBar().showMessage(f'已处理 {n} 个文件'))
        self.worker.finished.connect(lambda:(self.rule_table.setEnabled(True),self.template.setEnabled(True)))
        def finish(result):
            try:done(result)
            except Exception as e:self.bar.setRange(0,100);self.bar.setValue(0);self.statusBar().showMessage('处理返回，但界面更新失败，请保存现有项目并反馈');self.notice(str(e));return
            self.bar.setRange(0,100);self.bar.setValue(100)
        self.worker.done.connect(finish);self.worker.failed.connect(lambda s:(self.bar.setRange(0,100),self.bar.setValue(0),self.statusBar().showMessage('操作失败或取消，未完成'),self.notice(s)));self.worker.start()
    def start_scan(self):
        self.idle()
        if not self.project.people:raise ValueError('请先导入名单。')
        rules=self.get_rules();root=QFileDialog.getExistingDirectory(self,'选择材料文件夹')
        if not root:return
        candidate=copy.deepcopy(self.project)
        if candidate.rules!=rules:
            self.notice('规则变化只要求受影响的人工确认复核，其他确认保留。')
            update_rules(candidate,rules)
        def fn(cancel,progress):
            found,errors=scan(root,cancel,progress);merge_scan(candidate,root,found);return candidate,errors
        def done(result):
            self.project,errors=result;self.refresh();self.nav.setCurrentRow(2)
            if errors:self.notice('部分文件未扫描：\n'+'\n'.join(errors[:15]))
        self.run_job(fn,done)
    def results(self):
        l=self.page('2 · 核对结果','“已找到”只表示发现对应文件。重复和疑问请到人工确认页处理。')
        self.summary=QLabel('暂无结果');l.addWidget(self.summary);row=QHBoxLayout();l.addLayout(row)
        self.search=QLineEdit();self.search.setPlaceholderText('搜索姓名或编号');row.addWidget(self.search)
        self.filter=QComboBox();self.filter.addItems(['全部','缺件','已齐','重复']);row.addWidget(self.filter);self.search.textChanged.connect(self.refresh);self.filter.currentIndexChanged.connect(self.refresh)
        self.result_table=QTableWidget(0,5);self.result_table.setHorizontalHeaderLabels(['姓名','编号','已找到材料','缺件','异常']);self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch);self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers);l.addWidget(self.result_table)
        self.button(l,'查看当前人员详情',self.details);self.button(l,'补交材料：继续扫描新文件夹',lambda:self.guarded(self.start_scan));self.button(l,'取消当前扫描 / 导出',self.cancel)
    def details(self):
        i=self.result_table.currentRow()
        if i>=0:
            p=self.visible_people[i];mats,miss,dup=status(self.project,p);self.notice(p.name+'\n'+'\n'.join(f'{m.relative} · {m.category or "未分类"} · {m.reason}' for m in mats)+'\n缺件：'+'、'.join(miss))
    def manual_page(self):
        l=self.page('3 · 人工确认','人工确认优先于自动判断。忽略重复仅影响输出，不会删除原材料。')
        self.file_table=QTableWidget(0,4);self.file_table.setHorizontalHeaderLabels(['文件','归属','类别','匹配依据']);self.file_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch);self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers);l.addWidget(self.file_table)
        row=QHBoxLayout();l.addLayout(row);self.file_search=QLineEdit();self.file_search.setPlaceholderText('搜索文件名、归属或原因');row.addWidget(self.file_search)
        self.file_filter=QComboBox();self.file_filter.addItems(['全部文件','待确认','失效 / 需复核','重复','已忽略']);row.addWidget(self.file_filter)
        self.person_select=QComboBox();self.category_select=QComboBox();l.addWidget(QLabel('选择人员'));l.addWidget(self.person_select);l.addWidget(QLabel('选择材料类别'));l.addWidget(self.category_select)
        self.file_table.itemSelectionChanged.connect(self.reflect_selection)
        self.file_search.textChanged.connect(self.filter_files);self.file_filter.currentIndexChanged.connect(self.filter_files)
        self.button(l,'确认选中文件的归属和类别',lambda:self.guarded(self.manual_assign));self.button(l,'清除选中文件归属',lambda:self.guarded(lambda:self.manual_assign(clear=True)));self.button(l,'忽略 / 恢复选中文件（不删除）',lambda:self.guarded(lambda:self.manual_assign(ignore=True)))
    def manual_assign(self,clear=False,ignore=False):
        self.idle()
        i=self.file_table.currentRow()
        if i<0:raise ValueError('请先选择一个文件。')
        if ignore:
            m=self.project.materials[i];m.ignored=not m.ignored;self.project.event('切换忽略',path=m.path,ignored=m.ignored);self.refresh();return
        assign(self.project,i,None if clear else self.person_select.currentData(),None if clear else self.category_select.currentData(),ignore);self.refresh()
    def reflect_selection(self):
        i=self.file_table.currentRow()
        if 0<=i<len(self.project.materials):
            m=self.project.materials[i];self.person_select.setCurrentIndex(max(0,self.person_select.findData(m.person or m.previous_person)));self.category_select.setCurrentIndex(max(0,self.category_select.findData(m.category or m.previous_category)))
    def filter_files(self,*args):
        q=self.file_search.text().strip().casefold();mode=self.file_filter.currentText()
        duplicate_keys={(p.key,c) for p in self.project.people for c in status(self.project,p)[2]} if mode=='重复' else set()
        for i,m in enumerate(self.project.materials):
            text=' '.join(self.file_table.item(i,j).text() for j in range(4) if self.file_table.item(i,j)).casefold()
            duplicate=(m.person,m.category) in duplicate_keys
            hide=bool(q and q not in text) or (mode=='待确认' and bool(m.person and m.category) and not m.needs_review and not m.invalid and unchanged(m)) or (mode=='失效 / 需复核' and not (m.needs_review or m.invalid or not unchanged(m))) or (mode=='重复' and not duplicate) or (mode=='已忽略' and not m.ignored)
            self.file_table.setRowHidden(i,hide)
            if hide and self.file_table.currentRow()==i:self.file_table.clearSelection();self.file_table.setCurrentCell(-1,-1)
    def export_page(self):
        l=self.page('4 · 整理导出','复制到全新子文件夹，生成三份 Excel 和 ZIP。原材料不移动、不改名。')
        self.output=QLineEdit();l.addWidget(QLabel('选择与材料目录互不包含的输出父目录'));l.addWidget(self.output);self.button(l,'选择输出目录',lambda:self.output.setText(QFileDialog.getExistingDirectory(self,'选择输出目录')))
        self.preview=QPlainTextEdit();self.preview.setReadOnly(True);l.addWidget(self.preview);self.button(l,'预览整理路径',lambda:self.guarded(self.show_plan));self.button(l,'确认并复制 / 导出报告 / 打包',lambda:self.guarded(self.start_export));self.button(l,'取消当前操作',self.cancel)
    def show_plan(self):
        self.idle();self.sync_rules();self.refresh()
        if not self.output.text():raise ValueError('请选择输出目录。')
        plan=export_plan(self.project,self.output.text());self.preview.setPlainText(f'将复制 {len(plan)} 个文件，生成时间戳独立目录。\n'+'\n'.join(f'{m.relative} → {p}' for m,p in plan));return plan
    def start_export(self):
        self.idle();self.sync_rules();self.refresh()
        if not self.project.people:raise ValueError('尚未导入名单。')
        self.show_plan()
        if QMessageBox.question(self,'确认整理','待确认文件会放入“待确认文件”；重复材料会保留。继续复制并导出？')!=QMessageBox.Yes:return
        candidate=copy.deepcopy(self.project)
        output_parent=self.output.text()
        def done(path):
            self.project=candidate;self.notice('整理完成：\n'+str(path))
            try:os.startfile(path)
            except OSError:self.notice('导出完成，但系统未能打开文件夹。请手动打开：'+str(path))
        self.run_job(lambda c,p:export(candidate,output_parent,c,p),done)
    def history(self):
        l=self.page('历史与项目保存','项目文件包含姓名、编号、路径和人工确认。请保存在自己可信的位置。')
        self.button(l,'保存项目',lambda:self.guarded(self.save_project));self.button(l,'打开项目',self.open_project);self.button(l,'打开备份项目（.bak）',self.open_project);l.addStretch()
    def save_project(self):
        self.idle()
        p,_=QFileDialog.getSaveFileName(self,'保存项目',self.project_path or '核对任务.msq','核对项目 (*.msq)')
        if not p:return False
        self.sync_rules();self.refresh();save(self.project,p);self.project_path=p;self.saved_snapshot=self.state_snapshot();self.notice('项目已保存。原项目的上一版保留为 .bak。');return True
    def state_snapshot(self):
        return json.dumps({'project':asdict(self.project),'rules':[asdict(r) for r in self.get_rules()]},ensure_ascii=False,sort_keys=True)
    def open_project(self):
        if self.worker and self.worker.isRunning():self.notice('请先等待或取消当前操作。');return
        p,_=QFileDialog.getOpenFileName(self,'打开项目','','项目 (*.msq *.bak)')
        if not p:return
        def action():
            candidate=load(p)
            if not self.discard_confirm():return
            self.project=candidate;self.project_path=p;self.rule_table.setRowCount(0)
            for r in candidate.rules:self.add_rule(r)
            self.refresh();self.saved_snapshot=self.state_snapshot();self.nav.setCurrentRow(2)
        self.guarded(action)
    def help_page(self):
        l=self.page('使用帮助','无需登录，核心流程无需联网。')
        t=QTextBrowser();t.setPlainText('① 名单第一行放表头，例如姓名、学号；学号使用文本，避免前导零和长数字丢失。\n② 导入时预览并选择列。重复姓名或编号不会被当作可靠唯一身份。\n③ 每类材料填写扩展名，例如 .pdf,.docx；关键词可帮助区分类别。\n④ 选择材料目录，查看缺件。文件内容不会被审核。\n⑤ 人工确认页选择文件、人员和类别。忽略重复不删除原文件。\n⑥ 选择独立输出目录，先预览再导出。导出会复制文件并核验内容。\n⑦ 保存项目；补交时重新扫描新目录。同路径内容变化必须重新确认。\n\n安全与限制\n压缩包只当普通文件，不解压。链接目录与链接文件跳过。Excel读取活动工作表。无法扫描会显示原因。项目中保存的路径可能包含个人信息。源材料被移动后需要重新扫描。发生取消或失败会保留标记为未完成的输出，不能当正式交付。\n\n软件尚需通过发行验收。未签名程序可能显示 Windows 信誉提示；不应关闭系统防护。');l.addWidget(t)
    def refresh(self,*args):
        if not hasattr(self,'result_table'):return
        self.roster_label.setText(f'名单 {len(self.project.people)} 人 · 材料 {len(self.project.materials)} 个')
        complete=sum(not status(self.project,p)[1] for p in self.project.people);pending=sum(bool(not m.ignored and (not m.person or not m.category or m.needs_review or m.invalid or not unchanged(m))) for m in self.project.materials)
        self.summary.setText(f'共 {len(self.project.people)} 人　已找到全部必交文件 {complete} 人　缺件 {len(self.project.people)-complete} 人　待确认文件 {pending} 个\n依据上次扫描和当前文件大小/时间；导出前逐校验复制前后字节一致，不审核材料内容。')
        visible=[]
        for p in self.project.people:
            mats,miss,dup=status(self.project,p);q=self.search.text().strip()
            if q and q not in p.name and q not in p.identifier:continue
            mode=self.filter.currentText()
            if (mode=='缺件' and not miss) or (mode=='已齐' and miss) or (mode=='重复' and not dup):continue
            visible.append(p)
        self.visible_people=visible;self.result_table.setRowCount(len(visible))
        for i,p in enumerate(visible):
            mats,miss,dup=status(self.project,p)
            for j,v in enumerate([p.name,p.identifier,'、'.join(sorted({m.category for m in mats if m.category})),'、'.join(miss),'重复：'+'、'.join(dup) if dup else '']):self.result_table.setItem(i,j,QTableWidgetItem(v))
        names={p.key:p.name+' '+p.identifier for p in self.project.people};self.file_table.setRowCount(len(self.project.materials))
        for i,m in enumerate(self.project.materials):
            for j,v in enumerate([m.relative,names.get(m.person,'待确认'),m.category or '待确认',('已忽略 · ' if m.ignored else '')+('需复核 · ' if m.needs_review else '')+('源失效 · ' if m.invalid or not unchanged(m) else '')+m.reason]):self.file_table.setItem(i,j,QTableWidgetItem(v))
        self.person_select.clear();self.person_select.addItem('待确认',None)
        for p in self.project.people:self.person_select.addItem(p.name+' '+p.identifier,p.key)
        self.category_select.clear();self.category_select.addItem('待确认',None)
        for r in self.project.rules:self.category_select.addItem(r.name,r.name)
        self.reflect_selection()
        self.filter_files()
    def example(self):
        if self.worker and self.worker.isRunning():self.notice('请先等待或取消当前操作。');return
        if not self.discard_confirm():return
        from sample_data import make_sample
        root=Path(tempfile.mkdtemp(prefix='材料收齐示例_'));self.project=make_sample(root);self.project_path=None;self.saved_snapshot=None;self.rule_table.setRowCount(0)
        for r in self.project.rules:self.add_rule(r)
        self.refresh();self.nav.setCurrentRow(2)
    def cancel(self):
        if self.worker and self.worker.isRunning():self.worker.requestInterruption();self.statusBar().showMessage('已请求取消，等待当前文件安全处理结束')
    def closeEvent(self,e):
        if self.worker and self.worker.isRunning():self.cancel();self.notice('正在安全结束操作，请完成后再关闭。');e.ignore();return
        if not self.project.people:return
        try:dirty=self.state_snapshot()!=self.saved_snapshot
        except Exception:dirty=True
        if not dirty:return
        choice=QMessageBox.question(self,'保存后再退出？','当前任务有尚未保存的修改。保存项目不会备份原材料。',QMessageBox.Save|QMessageBox.Discard|QMessageBox.Cancel,QMessageBox.Save)
        if choice==QMessageBox.Cancel:e.ignore();return
        if choice==QMessageBox.Save:
            try:
                if not self.save_project():e.ignore()
            except Exception as error:self.notice('保存未完成，软件继续保持打开：\n'+str(error));e.ignore()

if __name__=='__main__':
    if '--restore-test' in sys.argv:
        os.environ['QT_QPA_PLATFORM']='offscreen'
        path=Path(sys.argv[sys.argv.index('--restore-test')+1]);output=Path(sys.argv[sys.argv.index('--restore-test')+2]);output.mkdir(parents=True,exist_ok=True)
        app=QApplication(sys.argv);w=Window();w.project=load(path);w.rule_table.setRowCount(0)
        for r in w.project.rules:w.add_rule(r)
        w.refresh();w.nav.setCurrentRow(2);w.show();app.processEvents()
        manual=sum(m.manual and bool(m.person) for m in w.project.materials)
        if manual!=1:raise RuntimeError('跨进程人工状态恢复失败')
        (output/'restore_selftest.json').write_text(json.dumps({'frozen':bool(getattr(sys,'frozen',False)),'manual_assignments_restored':manual,'people':len(w.project.people),'pid':os.getpid()},ensure_ascii=False),encoding='utf-8');w.hide();sys.exit(0)
    if '--self-test' in sys.argv:
        os.environ['QT_QPA_PLATFORM']='offscreen'
        import faulthandler
        output=Path(sys.argv[sys.argv.index('--self-test')+1]);output.mkdir(parents=True,exist_ok=True)
        trace=(output/'startup_trace.txt').open('w',encoding='utf-8');faulthandler.enable(file=trace);faulthandler.dump_traceback_later(15,file=trace)
        from PySide6.QtTest import QTest
        app=QApplication(sys.argv);w=Window();w.show();QTest.qWait(100);w.example();QTest.qWait(100)
        w.grab().save(str(output/'exe_results.png'))
        anon=next(i for i,m in enumerate(w.project.materials) if '匿名' in m.relative);assign(w.project,anon,w.project.people[-1].key,'照片')
        save(w.project,output/'sample.msq');restored=load(output/'sample.msq')
        result=export(restored,output/'export');report={'frozen':bool(getattr(sys,'frozen',False)),'people':len(restored.people),'files':len(restored.materials),'missing':sum(bool(status(restored,p)[1]) for p in restored.people),'export':str(result),'platform':sys.platform,'qt':'offscreen','executable':sys.executable}
        (output/'exe_selftest.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');faulthandler.cancel_dump_traceback_later();w.hide();sys.exit(0)
    app=QApplication(sys.argv);app.setApplicationName('材料收齐助手');w=Window();w.show();sys.exit(app.exec())

