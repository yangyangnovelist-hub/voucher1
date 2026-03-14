import flet as ft
import socket
import pandas as pd
import os, sys, re
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import company_manager as cm
from utils.subject import SubjectMatcher
from utils.excel_writer import write_voucher_excel
from utils.date_utils import extract_voucher_date, detect_date_col
from processor.ar import process_ar, DATE_COLS as AR_DATE_COLS
from processor.ap import process_ap, DATE_COLS as AP_DATE_COLS
from processor.bank import process_bank, read_bank_file, generate_pending_rows

def _excel_engine(path: str) -> str:
    """xls 用 xlrd，xlsx 用 calamine"""
    return "xlrd" if str(path).lower().endswith(".xls") else "calamine"


# ── 共用工具 ──────────────────────────────────────────────────

def extract_banks(m):
    if not m: return []
    result = []
    for code, name in sorted(m.code_to_name.items()):
        if not str(code).startswith("1002") or str(code) == "1002": continue
        parts = name.split("_")
        # 格式可能是: "银行存款_农商行_账号" 或 "农商行_账号" 或 "农商行账号"
        if len(parts) >= 3:
            bank_name = parts[-2].strip()
            last = parts[-1].strip()
            account_no = last if re.match(r"^\d+$", last) else ""
        elif len(parts) == 2:
            last = parts[-1].strip()
            mr = re.match(r"^([^\d]*)(\d+)?$", last)
            bank_name = (mr.group(1).strip() if mr else last)
            account_no = (mr.group(2) or "") if mr else ""
        else:
            mr = re.match(r"^([^\d]*)(\d+)?$", name)
            bank_name = (mr.group(1).strip() if mr else name)
            account_no = (mr.group(2) or "") if mr else ""
        result.append({"a": account_no, "b": bank_name, "s": str(code)})
    return result


def mask_account(acct: str) -> str:
    """账号脱敏：保留前4后4，中间替换为****"""
    if not acct or len(acct) <= 8:
        return acct
    return acct[:4] + "****" + acct[-4:]

def bank_subj_opts(m):
    if not m: return {}
    return {f"{c}  {n}": c for c, n in m.code_to_name.items()
            if str(c).startswith("1002") and str(c) != "1002"}


# ── 银行账号编辑器组件 ────────────────────────────────────────

class BankEditor(ft.Column):
    def __init__(self, rows_data, subj_opts, page):
        super().__init__(spacing=6)
        self.rows_data = rows_data
        self.subj_opts = subj_opts
        self._pg = page
        self.rebuild()

    def rebuild(self):
        self.controls.clear()
        self.controls.append(ft.Row([
            ft.Text("账号 *",       width=200, size=12, weight=ft.FontWeight.W_500),
            ft.Text("银行名称",      width=140, size=12, weight=ft.FontWeight.W_500),
            ft.Text("银行存款科目 *", width=240, size=12, weight=ft.FontWeight.W_500),
        ], spacing=8))
        for i, row in enumerate(self.rows_data):
            self.controls.append(self._make_row(i, row))
        self.controls.append(
            ft.TextButton("＋ 添加账号", on_click=self._add_row, icon=ft.Icons.ADD)
        )
        empty_acct = [r for r in self.rows_data if r.get("s") and not r.get("a","").strip()]
        if empty_acct:
            self.controls.append(
                ft.Text(f"⚠️ {len(empty_acct)} 个账号未填写，保存后这些行将被忽略",
                       color=ft.Colors.RED_600, size=12)
            )
        if self._pg: self._pg.update()

    def _make_row(self, idx, row):
        def on_acct(e, i=idx):  self.rows_data[i]["a"] = e.control.value
        def on_name(e, i=idx):  self.rows_data[i]["b"] = e.control.value
        def on_code(e, i=idx):  self.rows_data[i]["s"] = self.subj_opts.get(e.control.value, e.control.value)

        acct_f = ft.TextField(value=row.get("a",""), hint_text="完整账号", width=200,
                              on_blur=on_acct, dense=True)
        name_f = ft.TextField(value=row.get("b",""), hint_text="如：招商银行", width=140,
                              on_blur=on_name, dense=True)

        if self.subj_opts:
            opts  = list(self.subj_opts.keys())
            cur   = next((k for k,v in self.subj_opts.items() if v == row.get("s","")), opts[0] if opts else "")
            subj_f = ft.Dropdown(value=cur,
                                 options=[ft.dropdown.Option(key=k, text=k) for k in opts],
                                 width=240, dense=True, enable_filter=True, enable_search=True, on_select=on_code)
        else:
            subj_f = ft.TextField(value=row.get("s",""), hint_text="如：0000001",
                                  width=240, on_blur=on_code, dense=True)

        def up(e, i=idx):
            if i > 0:
                self.rows_data[i], self.rows_data[i-1] = self.rows_data[i-1], self.rows_data[i]
                self.rebuild()
        def dn(e, i=idx):
            if i < len(self.rows_data)-1:
                self.rows_data[i], self.rows_data[i+1] = self.rows_data[i+1], self.rows_data[i]
                self.rebuild()
        def dl(e, i=idx):
            if len(self.rows_data) > 1:
                self.rows_data.pop(i)
                self.rebuild()

        return ft.Row([
            acct_f, name_f, subj_f,
            ft.IconButton(ft.Icons.ARROW_UPWARD,   on_click=up, icon_size=18, disabled=(idx==0)),
            ft.IconButton(ft.Icons.ARROW_DOWNWARD,  on_click=dn, icon_size=18,
                          disabled=(idx==len(self.rows_data)-1)),
            ft.IconButton(ft.Icons.DELETE_OUTLINE,  on_click=dl, icon_size=18,
                          icon_color=ft.Colors.RED_400, disabled=(len(self.rows_data)<=1)),
        ], spacing=6)

    def _add_row(self, e):
        self.rows_data.append({"a":"","b":"","s":""})
        self.rebuild()

    def get_accounts(self):
        return [cm.BankAccount(account_no=r["a"].strip(),
                               bank_name=r["b"].strip(),
                               subject_code=r["s"].strip())
                for r in self.rows_data if r.get("a","").strip()]


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main(page: ft.Page):
    page.title = "凭证生成工具"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window.width  = 1100
    page.window.height = 780
    page.window.min_width  = 900
    page.window.min_height = 600
    page.padding = 0
    page.fonts   = {"default": "Microsoft YaHei"}

    body = ft.Column(expand=True, scroll=ft.ScrollMode.AUTO, spacing=0)
    page.add(ft.Container(body, expand=True))

    def navigate(controls):
        body.controls.clear()
        body.controls.extend(controls if isinstance(controls, list) else [controls])
        page.update()

    def snack(msg, color=ft.Colors.GREEN_700):
        page.show_dialog(
            ft.SnackBar(ft.Text(msg, color=ft.Colors.WHITE), bgcolor=color)
        )

    # ── 公司列表 ────────────────────────────────────────────
    def show_list():
        companies = cm.list_companies()
        cards = []
        for sn in companies:
            p = cm.get_company(sn)
            ok = bool(p.subject_file and os.path.exists(os.path.join(cm.DATA_DIR, p.subject_file)))
            def click(e, key=sn): show_workspace(key)
            cards.append(ft.Card(ft.Container(
                ft.Column([
                    ft.Row([ft.Icon(ft.Icons.BUSINESS,
                                   color=ft.Colors.BLUE_600 if ok else ft.Colors.ORANGE_400),
                            ft.Text(sn, weight=ft.FontWeight.BOLD)]),
                    ft.Text(p.name, size=12, color=ft.Colors.GREY_700,
                            max_lines=2, overflow=ft.TextOverflow.ELLIPSIS),
                    ft.Text("✅ 科目已上传" if ok else "⚠️ 待上传科目",
                            size=11, color=ft.Colors.GREEN_700 if ok else ft.Colors.ORANGE_700),
                ], spacing=8),
                padding=16, width=210, on_click=click, ink=True,
            )))

        cards.append(ft.Card(ft.Container(
            ft.Column([
                ft.Icon(ft.Icons.ADD_CIRCLE_OUTLINE, size=36, color=ft.Colors.BLUE_400),
                ft.Text("添加新公司", weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_600),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8),
            padding=16, width=210, height=110,
            on_click=lambda e: show_add(), ink=True,
            alignment=ft.Alignment(0, 0),
        )))

        navigate([
            ft.Container(ft.Column([
                ft.Text("📊 凭证生成工具", size=26, weight=ft.FontWeight.BOLD,
                        color=ft.Colors.BLUE_800),
                ft.Text("选择公司开始操作", size=13, color=ft.Colors.GREY_600),
            ]), padding=ft.padding.only(left=24, top=24, bottom=20)),
            ft.Container(
                ft.Row(cards, wrap=True, spacing=14, run_spacing=14),
                padding=ft.padding.symmetric(horizontal=24),
            ),
        ])

    # ── 添加新公司 ──────────────────────────────────────────
    def show_add():
        short_f   = ft.TextField(label="公司简称 *", hint_text="如：XX公司", width=260)
        full_f    = ft.TextField(label="公司全称（户名）*",
                                 hint_text="如：XX贸易有限公司", width=480)
        subj_txt  = ft.Text("未上传", color=ft.Colors.GREY_500, size=12)
        subj_path = {"v": None}
        bank_data = [{"a":"","b":"","s":""}]
        sopts     = {}
        editor_wrap = ft.Column()

        def rebuild_editor():
            editor_wrap.controls.clear()
            editor_wrap.controls.append(BankEditor(bank_data, sopts, page))
            page.update()

        subj_picker = ft.FilePicker()
        page.services.append(subj_picker)
        rebuild_editor()

        async def pick_subj(e):
            files = await subj_picker.pick_files(allowed_extensions=["xls","xlsx"])
            if not files: return
            p = files[0].path
            subj_path["v"] = p
            try:
                m = SubjectMatcher(pd.read_excel(p, dtype=str, engine=_excel_engine(p)))
                sopts.clear(); sopts.update(bank_subj_opts(m))
                extracted = extract_banks(m)
                if extracted:
                    bank_data.clear(); bank_data.extend(extracted)
                    subj_txt.value = f"✅ 已上传，识别到 {len(extracted)} 个银行账号"
                else:
                    subj_txt.value = "✅ 已上传"
                subj_txt.color = ft.Colors.GREEN_700
                rebuild_editor()
            except Exception as ex:
                subj_txt.value = f"读取失败: {ex}"; subj_txt.color = ft.Colors.RED_600
            page.update()

        def save(e):
            short = short_f.value.strip(); full = full_f.value.strip()
            if not short or not full:
                snack("简称和全称不能为空", ft.Colors.RED_700); return
            if short in cm.list_companies():
                snack(f"「{short}」已存在", ft.Colors.RED_700); return
            editor = editor_wrap.controls[0]
            accts = editor.get_accounts()
            prof = cm.CompanyProfile(name=full, short_name=short, bank_accounts=accts)
            sb = open(subj_path["v"],"rb").read() if subj_path["v"] else None
            cm.save_company(prof, sb)
            page.services.remove(subj_picker)
            snack(f"✅ 已保存「{short}」")
            show_list()

        def go_back(e):
            page.services.remove(subj_picker)
            show_list()

        navigate([
            ft.Container(ft.Row([
                ft.IconButton(ft.Icons.ARROW_BACK, on_click=go_back),
                ft.Text("添加新公司", size=20, weight=ft.FontWeight.BOLD),
            ]), padding=ft.padding.only(left=12, top=16, bottom=8)),
            ft.Container(ft.Column([
                ft.Row([short_f, full_f], spacing=16),
                ft.Divider(),
                ft.Text("科目列表", weight=ft.FontWeight.W_600),
                ft.Row([
                    ft.ElevatedButton("选择科目文件", icon=ft.Icons.UPLOAD_FILE,
                                      on_click=pick_subj),
                    subj_txt,
                ], spacing=12),
                ft.Divider(),
                ft.Text("银行账号", weight=ft.FontWeight.W_600),
                ft.Text("列表顺序即优先顺序；账号是流水筛选唯一依据",
                        size=12, color=ft.Colors.GREY_600),
                editor_wrap,
                ft.Divider(),
                ft.ElevatedButton("💾 保存并进入", on_click=save, icon=ft.Icons.SAVE,
                                  style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_600,
                                                       color=ft.Colors.WHITE)),
            ], spacing=12), padding=24),
        ])

    # ── 公司工作区 ──────────────────────────────────────────
    def show_workspace(sn):
        profile = cm.get_company(sn)
        matcher = cm.load_matcher(sn)

        gen_tab = build_gen_tab(sn, profile, matcher)
        cfg_tab = build_cfg_tab(sn, profile, matcher)

        navigate([
            ft.Container(ft.Row([
                ft.Icon(ft.Icons.FOLDER_OPEN, color=ft.Colors.BLUE_600),
                ft.Text(f"{profile.short_name}　{profile.name}",
                        size=17, weight=ft.FontWeight.BOLD),
                ft.Container(expand=True),
                ft.OutlinedButton("切换公司", icon=ft.Icons.SWAP_HORIZ,
                                  on_click=lambda e: show_list()),
            ]), padding=ft.padding.only(left=20, right=20, top=14, bottom=6)),
            ft.Divider(height=1),
            ft.Tabs(
                selected_index=0,
                length=2,
                content=ft.Column([
                    ft.TabBar(tabs=[
                        ft.Tab(label="🧾 生成凭证"),
                        ft.Tab(label="⚙️ 公司设置"),
                    ]),
                    ft.TabBarView(controls=[
                        ft.Container(gen_tab, padding=20),
                        ft.Container(cfg_tab, padding=20),
                    ], expand=True),
                ], expand=True, spacing=0),
                expand=True,
            ),
        ])

    # ── 生成凭证 Tab ────────────────────────────────────────
    def build_gen_tab(sn, profile, matcher):
        if not matcher:
            return ft.Container(
                ft.Column([
                    ft.Icon(ft.Icons.WARNING_AMBER, size=48, color=ft.Colors.ORANGE_400),
                    ft.Text("尚未上传科目列表，请切换到「公司设置」标签上传。", size=15),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                alignment=ft.Alignment(0, 0), padding=60,
            )

        file_bytes = {"ar": None, "ap": None, "bank": None}
        ar_lbl    = ft.Text("未选择", size=12, color=ft.Colors.GREY_500)
        ap_lbl    = ft.Text("未选择", size=12, color=ft.Colors.GREY_500)
        bank_lbl  = ft.Text("未选择", size=12, color=ft.Colors.GREY_500)
        bank_rule_lbl = ft.Text("", size=11, color=ft.Colors.GREY_600)

        ar_picker   = ft.FilePicker()
        ap_picker   = ft.FilePicker()
        bank_picker = ft.FilePicker()
        save_picker = ft.FilePicker()
        page.services.extend([ar_picker, ap_picker, bank_picker, save_picker])

        def update_bank_rule(acct_no: str = ""):
            accts = profile.bank_accounts if profile else []
            if not accts:
                bank_rule_lbl.value = ""
                return
            order = " > ".join([mask_account(a.account_no) for a in accts if a.account_no])
            idx = None
            if acct_no:
                for i, a in enumerate(accts, 1):
                    if a.account_no.strip() == acct_no.strip():
                        idx = i
                        break
            if idx is None:
                idx = 1
            bank_rule_lbl.value = (
                f"内部转账去重：按银行账号列表顺序生效；当前账号序号 {idx}/{len(accts)}。"
                f"第1个账号保留全部内部转账，其余账号过滤与更早账号之间的转账。"
                f"  账号顺序：{order}"
            )

        async def pick_ar(e):
            files = await ar_picker.pick_files(allowed_extensions=["xls","xlsx"])
            if files:
                file_bytes["ar"] = open(files[0].path,"rb").read()
                ar_lbl.value = f"✅ {files[0].name}"
                ar_lbl.color = ft.Colors.GREEN_700; page.update()

        async def pick_ap(e):
            files = await ap_picker.pick_files(allowed_extensions=["xls","xlsx"])
            if files:
                file_bytes["ap"] = open(files[0].path,"rb").read()
                ap_lbl.value = f"✅ {files[0].name}"
                ap_lbl.color = ft.Colors.GREEN_700; page.update()

        async def pick_bank(e):
            files = await bank_picker.pick_files(allowed_extensions=["xls","xlsx"])
            if files:
                raw = open(files[0].path,"rb").read()
                file_bytes["bank"] = raw
                try:
                    _, acct_no, _ = read_bank_file(raw)
                    m = cm.find_bank_by_account(sn, acct_no)
                    bank_lbl.value = (f"✅ {files[0].name}　({acct_no} ✅)" if m
                                      else f"✅ {files[0].name}　({acct_no} 未登记)")
                    update_bank_rule(acct_no)
                except:
                    bank_lbl.value = f"✅ {files[0].name}"
                    update_bank_rule("")
                bank_lbl.color = ft.Colors.GREEN_700; page.update()

        which_ar   = ft.Checkbox(label="应收凭证",    value=True)
        which_ap   = ft.Checkbox(label="应付凭证",    value=True)
        which_bank = ft.Checkbox(label="银行收支凭证", value=True)
        ar_no  = ft.TextField(label="应收凭证号",  value="1", width=120, keyboard_type=ft.KeyboardType.NUMBER)
        ap_no  = ft.TextField(label="应付凭证号",  value="1", width=120, keyboard_type=ft.KeyboardType.NUMBER)
        b_inc  = ft.TextField(label="银行收入号",  value="1", width=120, keyboard_type=ft.KeyboardType.NUMBER)
        b_exp  = ft.TextField(label="银行支出号",  value="1", width=120, keyboard_type=ft.KeyboardType.NUMBER)
        # 应付贷方改为固定走“挂账”逻辑（按设计不再暴露选择项）

        update_bank_rule("")

        progress    = ft.ProgressRing(visible=False, width=32, height=32)
        results_col = ft.Column(visible=False, spacing=10)
        output_cache = {}

        async def do_save(e):
            path = await save_picker.get_directory_path(dialog_title="选择保存位置")
            if path:
                for fn, buf in output_cache.items():
                    with open(os.path.join(path, fn), "wb") as f:
                        f.write(buf.getvalue())
                snack(f"✅ 已保存 {len(output_cache)} 个文件到：{path}")

        async def generate(e):
            import asyncio
            progress.visible = True; results_col.visible = False; page.update()
            out, warns, pending, ok = {}, [], [], True
            output_cache.clear()

            _ar  = int(ar_no.value  or 1)
            _ap  = int(ap_no.value  or 1)
            _inc = int(b_inc.value  or 1)
            _exp = int(b_exp.value  or 1)
            if which_ar.value and not file_bytes["ar"]:
                snack("请先选择应收发票文件", ft.Colors.RED_700)
                progress.visible = False; page.update(); return
            if which_ap.value and not file_bytes["ap"]:
                snack("请先选择应付发票文件", ft.Colors.RED_700)
                progress.visible = False; page.update(); return
            if which_bank.value and not file_bytes["bank"]:
                snack("请先选择银行流水文件", ft.Colors.RED_700)
                progress.visible = False; page.update(); return

            def _heavy():
                _out, _warns, _pend = {}, [], []
                if which_ar.value:
                    df = pd.read_excel(BytesIO(file_bytes["ar"]),
                                       dtype={"数电发票号码":str,"发票号码":str}, engine="calamine")
                    rows, w = process_ar(df, matcher, _ar)
                    _warns += w; _out["凭证_应收.xlsx"] = write_voucher_excel(rows,"应收")
                if which_ap.value:
                    df = pd.read_excel(BytesIO(file_bytes["ap"]),
                                       dtype={"数电发票号码":str,"发票号码":str}, engine="calamine")
                    rows, w = process_ap(df, matcher, _ap, 'pending')
                    _warns += w; _out["凭证_应付.xlsx"] = write_voucher_excel(rows,"应付")
                if which_bank.value:
                    df_bank, facct, fco = read_bank_file(file_bytes["bank"])
                    acct_obj = cm.find_bank_by_account(sn, facct)
                    if not acct_obj and profile.bank_accounts:
                        acct_obj = profile.bank_accounts[0]
                    result, w, pend = process_bank(
                        df_bank, matcher,
                        company_name=profile.name or fco,
                        income_voucher_start=_inc, expense_voucher_start=_exp,
                        bank_account_no=facct or (acct_obj.account_no if acct_obj else ""),
                        bank_subject_code=acct_obj.subject_code if acct_obj else "",
                        user_rules=cm.load_rules(sn),
                        company_accounts=profile.bank_accounts if profile else [],
                    )
                    _warns += w; _pend += pend
                    lbl = (f"{acct_obj.bank_name}{facct or acct_obj.account_no}"
                           if (acct_obj and acct_obj.bank_name)
                           else (facct or (acct_obj.account_no if acct_obj else "")))
                    for month, (ir, er) in result.items():
                        yr,mo = month.split("-")
                        if ir: _out[f"银行_{lbl}_{yr}年{mo}月_收入.xlsx"] = write_voucher_excel(ir,"收入")
                        if er: _out[f"银行_{lbl}_{yr}年{mo}月_支出.xlsx"] = write_voucher_excel(er,"支出")
                return _out, _warns, _pend

            try:
                out, warns, pending = await asyncio.to_thread(_heavy)
            except Exception as ex:
                import traceback; traceback.print_exc()
                snack(f"生成失败: {ex}", ft.Colors.RED_700); ok=False

            progress.visible = False
            if ok and (out or warns or pending):
                output_cache.clear(); output_cache.update(out)
                rc = []
                if out:
                    rc += [
                        ft.Text(f"生成完成：{len(out)} 个文件",
                                color=ft.Colors.GREEN_700, weight=ft.FontWeight.BOLD),
                        ft.ElevatedButton(
                            "💾 保存全部文件",
                            icon=ft.Icons.SAVE_ALT,
                            on_click=do_save,
                            style=ft.ButtonStyle(bgcolor=ft.Colors.GREEN_600, color=ft.Colors.WHITE),
                        ),
                    ]
                else:
                    rc.append(ft.Text("未生成可直接导出的凭证文件（存在待确认项目）",
                                      color=ft.Colors.ORANGE_700, weight=ft.FontWeight.BOLD))
                if warns:
                    rc.append(ft.ExpansionTile(
                        title=ft.Text(f"⚠️ {len(warns)} 条标黄提示（已在 Excel 中标注）",
                                      color=ft.Colors.ORANGE_700),
                        controls=[ft.Text(f"• {w}", size=12) for w in warns],
                    ))
                if pending:
                    rc += build_pending_ui(pending, matcher, output_cache,
                                           _inc, _exp, save_picker, sn)
                if out and not warns and not pending:
                    rc.append(ft.Text("🎉 全部自动匹配，无需手动确认！",
                                      color=ft.Colors.GREEN_700))
                results_col.controls.clear()
                results_col.controls.extend(rc)
                results_col.visible = True
            page.update()

        return ft.Column([
            ft.Text("① 上传源文件", size=15, weight=ft.FontWeight.W_600),
            ft.Row([
                ft.Column([
                    ft.ElevatedButton("选择应收发票", icon=ft.Icons.UPLOAD_FILE, on_click=pick_ar),
                    ar_lbl,
                ], spacing=4),
                ft.Column([
                    ft.ElevatedButton("选择应付发票", icon=ft.Icons.UPLOAD_FILE, on_click=pick_ap),
                    ap_lbl,
                ], spacing=4),
                ft.Column([
                    ft.ElevatedButton("选择银行流水", icon=ft.Icons.UPLOAD_FILE, on_click=pick_bank),
                    bank_lbl,
                    bank_rule_lbl,
                ], spacing=4),
            ], spacing=24),
            ft.Divider(),
            ft.Text("② 凭证参数", size=15, weight=ft.FontWeight.W_600),
            ft.Row([which_ar, which_ap, which_bank], spacing=20),
            ft.Row([ar_no, ap_no, b_inc, b_exp], spacing=12, wrap=True),
            ft.Divider(),
            ft.ElevatedButton(
                "🚀 生成凭证", on_click=generate, icon=ft.Icons.PLAY_ARROW,
                style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE),
                width=180, height=46,
            ),
            progress,
            ft.Divider(visible=False),
            results_col,
        ], spacing=12, scroll=ft.ScrollMode.AUTO, expand=True)

    def build_pending_ui(pending, matcher, output_cache, b_inc, b_exp, save_picker, sn):
        n_inc = sum(1 for p in pending if p.get("direction")=="income")
        n_exp = len(pending)-n_inc
        parts = []
        if n_inc: parts.append(f"{n_inc} 笔收入（贷方待定）")
        if n_exp: parts.append(f"{n_exp} 笔支出（借方待定）")

        all_opts = {f"{c}  {n}": c for c,n in sorted(matcher.code_to_name.items()) if n}
        assignments = {}
        save_rules_flags = {}
        note_kws = {}
        sum_kws  = {}

        item_controls = []
        for i, item in enumerate(pending):
            direction = item.get("direction","expense")
            dir_label = "请选贷方科目（收入）" if direction=="income" else "请选借方科目（支出）"
            # BUG FIX: on_select→on_change（on_select在此版本不触发赋值）
            # BUG FIX: editable=True 确保下拉列表可展开和搜索，不再只显示「跳过」
            dd = ft.Dropdown(
                label=dir_label, width=380,
                hint_text="点击展开 / 输入关键词搜索科目",
                editable=True,
                enable_filter=True,
                options=[ft.dropdown.Option(key="__skip__", text="（跳过）")] +
                        [ft.dropdown.Option(key=k, text=k) for k in list(all_opts.keys())],
                on_select=lambda e, idx=i: assignments.update(
                    {idx: all_opts.get(e.control.value, "")})
                    if e.control.value and e.control.value != "__skip__" else None,
            )
            save_cb = ft.Checkbox(label="保存为规则", value=False)
            nkw_f = ft.TextField(label="备注关键词", value=item.get("note",""),
                                  width=180, dense=True, visible=False,
                                  on_blur=lambda e, idx=i: note_kws.update({idx: e.control.value}))
            skw_f = ft.TextField(label="摘要关键词", value=item.get("summary",""),
                                  width=180, dense=True, visible=False,
                                  on_blur=lambda e, idx=i: sum_kws.update({idx: e.control.value}))
            rule_row = ft.Row([nkw_f, skw_f], spacing=8, visible=False)

            def toggle_rule(e, idx=i, rr=rule_row):
                save_rules_flags.update({idx: e.control.value})
                rr.visible = e.control.value
                page.update()
            save_cb.on_change = toggle_rule

            item_controls.append(ft.Card(ft.Container(ft.Column([
                ft.Row([
                    ft.Text(f"{'📥' if direction=='income' else '📤'} "
                            f"{item['trade_date']}  ¥{item['amount']:,.2f}  "
                            f"{item['counterpart'] or ''}",
                            weight=ft.FontWeight.W_500),
                ]),
                ft.Text(
                    (f"对方账号: {item['counterpart_acct']}  " if item.get('counterpart_acct') else "") +
                    f"摘要: {item['summary'] or '—'}  备注: {item['note'] or '—'}",
                    size=12, color=ft.Colors.GREY_600),
                dd, save_cb, rule_row,
            ], spacing=6), padding=12)))

        async def apply_pending(e):
            if not assignments:
                snack("请至少指定一笔的科目", ft.Colors.ORANGE_400); return
            for idx, flag in save_rules_flags.items():
                if flag and idx in assignments and assignments[idx]:
                    nkw = note_kws.get(idx, pending[idx].get("note","")).strip()
                    skw = sum_kws.get(idx,  pending[idx].get("summary","")).strip()
                    if nkw or skw:
                        cm.add_rule(sn, {
                            "match_note": nkw, "match_summary": skw,
                            "match_counterpart": "",
                            "subject_code": assignments[idx],
                            "label": (pending[idx].get("note") or pending[idx].get("summary",""))[:20],
                        })
            extra = generate_pending_rows(pending, assignments, matcher,
                                          income_voucher_start=b_inc,
                                          expense_voucher_start=b_exp)
            import openpyxl
            from utils.excel_writer import OUTPUT_COLS
            for month, rows in extra.items():
                yr, mo = month.split("-")
                inc_extra, exp_extra = [], []
                # BUG FIX: 只取当前月份的条目，避免跨月索引错位
                assigned_items = [
                    (i, p) for i, p in enumerate(pending)
                    if i in assignments and assignments[i] and p["month"] == month
                ]
                row_idx = 0
                for i, item in assigned_items:
                    pair = rows[row_idx:row_idx+2]
                    if item.get("direction") == "income": inc_extra.extend(pair)
                    else: exp_extra.extend(pair)
                    row_idx += 2

                for extra_rows, suffix in [(inc_extra, "收入"), (exp_extra, "支出")]:
                    if not extra_rows: continue
                    target = next((fn for fn in output_cache if f"{yr}年{mo}月_{suffix}" in fn), None)
                    if target:
                        buf = output_cache[target]; buf.seek(0)
                        wb = openpyxl.load_workbook(buf); ws = wb.active
                        existing = []
                        for r in ws.iter_rows(min_row=2, values_only=True):
                            rd = {OUTPUT_COLS[ci]: v for ci, v in enumerate(r) if ci < len(OUTPUT_COLS)}
                            rd["_yellow"] = False; existing.append(rd)
                        # BUG FIX: 合并后重新顺序编号，避免分录序号与已有条目冲突
                        merged = existing + extra_rows
                        for seq_idx, merged_row in enumerate(merged):
                            merged_row["分录序号"] = seq_idx + 1
                        output_cache[target] = write_voucher_excel(merged, suffix)
                    else:
                        # BUG FIX: 新建文件也要保证序号从1开始连续
                        for seq_idx, er in enumerate(extra_rows):
                            er["分录序号"] = seq_idx + 1
                        sample = next((f for f in output_cache if "银行_" in f), "")
                        prefix = sample.split(f"_{yr}")[0] if sample else "银行"
                        output_cache[f"{prefix}_{yr}年{mo}月_{suffix}.xlsx"] = \
                            write_voucher_excel(extra_rows, suffix)

            path = await save_picker.get_directory_path(dialog_title="选择保存位置")
            if path:
                for fn, buf in output_cache.items():
                    with open(os.path.join(path, fn), "wb") as f:
                        f.write(buf.getvalue())
                snack(f"✅ 已保存 {len(output_cache)} 个文件到：{path}")

        return [
            ft.Divider(),
            ft.Text(f"📋 {len(pending)} 笔待确认科目  —  " + "、".join(parts),
                    size=15, weight=ft.FontWeight.BOLD),
            ft.Text("指定科目后点击「确认并保存」，可选勾选「保存为规则」下次自动匹配",
                    size=12, color=ft.Colors.GREY_600),
            *item_controls,
            ft.ElevatedButton("✅ 确认并保存", on_click=apply_pending,
                              icon=ft.Icons.CHECK_CIRCLE_OUTLINE),
        ]

    # ── 公司设置 Tab ────────────────────────────────────────
    def build_cfg_tab(sn, profile, matcher):
        name_f = ft.TextField(label="公司全称（户名）", value=profile.name, width=500)
        has_subj = bool(profile.subject_file and
                        os.path.exists(os.path.join(cm.DATA_DIR, profile.subject_file)))
        subj_lbl = ft.Text("当前：" + ("已上传 ✅" if has_subj else "未上传 ⚠️"),
                           color=ft.Colors.GREEN_700 if has_subj else ft.Colors.ORANGE_700,
                           size=12)
        new_subj = {"v": None}
        sopts = bank_subj_opts(matcher)
        bank_data = [{"a":a.account_no,"b":a.bank_name,"s":a.subject_code}
                     for a in profile.bank_accounts] or [{"a":"","b":"","s":""}]
        editor_wrap = ft.Column()

        def rebuild():
            editor_wrap.controls.clear()
            editor_wrap.controls.append(BankEditor(bank_data, sopts, page))
            page.update()

        subj_picker = ft.FilePicker()
        page.services.append(subj_picker)

        async def pick_subj(e):
            files = await subj_picker.pick_files(allowed_extensions=["xls","xlsx"])
            if not files: return
            new_subj["v"] = files[0].path
            try:
                m2 = SubjectMatcher(pd.read_excel(files[0].path, dtype=str, engine=_excel_engine(files[0].path)))
                sopts.clear(); sopts.update(bank_subj_opts(m2))
                subj_lbl.value = "新文件已选 ✅"; subj_lbl.color = ft.Colors.GREEN_700
                extracted = extract_banks(m2)
                if extracted:
                    snack(f"识别到 {len(extracted)} 个银行账号，点「用新科目覆盖银行账号」按钮更新")
                page.update()
            except Exception as ex:
                subj_lbl.value = f"读取失败: {ex}"; subj_lbl.color = ft.Colors.RED_600
                page.update()

        def overwrite_banks(e):
            if not new_subj["v"]: return
            m2 = SubjectMatcher(pd.read_excel(new_subj["v"], dtype=str, engine=_excel_engine(new_subj["v"])))
            extracted = extract_banks(m2)
            if extracted:
                bank_data.clear(); bank_data.extend(extracted); rebuild()
                snack(f"已覆盖，识别到 {len(extracted)} 个账号")

        def save(e):
            editor = editor_wrap.controls[0]
            accts = editor.get_accounts()
            updated = cm.CompanyProfile(name=name_f.value, short_name=sn,
                                        bank_accounts=accts, subject_file=profile.subject_file)
            sb = open(new_subj["v"],"rb").read() if new_subj["v"] else None
            cm.save_company(updated, sb)
            page.services.remove(subj_picker)
            snack("✅ 已保存")
            show_workspace(sn)

        def confirm_delete(e):
            def do_del(e):
                page.pop_dialog()
                cm.delete_company(sn)
                page.services.remove(subj_picker)
                show_list()
                snack(f"已删除「{sn}」")

            def cancel(e):
                page.pop_dialog()

            dlg = ft.AlertDialog(
                title=ft.Text("确认删除"),
                content=ft.Text(f"确定删除「{sn}」及其所有数据？此操作不可恢复。"),
                actions=[
                    ft.TextButton("取消", on_click=cancel),
                    ft.ElevatedButton("删除", on_click=do_del,
                                      style=ft.ButtonStyle(bgcolor=ft.Colors.RED_600,
                                                           color=ft.Colors.WHITE)),
                ],
            )
            page.show_dialog(dlg)

        rules = cm.load_rules(sn)
        rules_col = ft.Column(spacing=4)

        def rebuild_rules():
            rules_col.controls.clear()
            if rules:
                rules_col.controls.append(
                    ft.Text(f"匹配规则 — {len(rules)} 条", weight=ft.FontWeight.W_600))
                for ri, rule in enumerate(rules):
                    conds = []
                    if rule.get("match_note"):        conds.append(f"备注含「{rule['match_note']}」")
                    if rule.get("match_summary"):     conds.append(f"摘要含「{rule['match_summary']}」")
                    if rule.get("match_counterpart"): conds.append(f"对方含「{rule['match_counterpart']}」")
                    def del_rule(e, r=ri):
                        rules.pop(r); cm.save_rules(sn, rules); rebuild_rules(); page.update()
                    rules_col.controls.append(ft.Row([
                        ft.Text("  +  ".join(conds)+f"  →  {rule['subject_code']}",
                                size=12, expand=True),
                        ft.IconButton(ft.Icons.DELETE_OUTLINE, on_click=del_rule, icon_size=16),
                    ]))
            page.update()

        rebuild_rules()
        rebuild()

        return ft.Column([
            name_f,
            ft.Divider(),
            ft.Text("科目列表", weight=ft.FontWeight.W_600),
            ft.Row([
                ft.ElevatedButton("选择新科目文件", icon=ft.Icons.UPLOAD_FILE,
                                  on_click=pick_subj),
                subj_lbl,
                ft.TextButton("用新科目覆盖银行账号", on_click=overwrite_banks),
            ], spacing=12),
            ft.Divider(),
            ft.Text("银行账号", weight=ft.FontWeight.W_600),
            editor_wrap,
            ft.Divider(),
            rules_col,
            ft.Divider(),
            ft.Row([
                ft.ElevatedButton("💾 保存设置", on_click=save, icon=ft.Icons.SAVE,
                                  style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_600,
                                                       color=ft.Colors.WHITE)),
                ft.OutlinedButton("🗑️ 删除公司", on_click=confirm_delete,
                                  style=ft.ButtonStyle(color=ft.Colors.RED_600)),
            ], spacing=16),
        ], spacing=12, scroll=ft.ScrollMode.AUTO, expand=True)

    # ── 启动 ────────────────────────────────────────────────
    show_list()


def _ensure_single_instance():
    """
    Prevent launching multiple instances (especially for Windows exe).
    Uses a localhost TCP bind as a simple single-instance lock.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 47299))
        s.listen(1)
        return s
    except OSError:
        return None


if __name__ == "__main__":
    _lock_sock = _ensure_single_instance()
    if _lock_sock is None:
        print("应用已在运行中，请切换到已打开的窗口。")
        raise SystemExit(0)

    # Flet API compatibility: newer versions use ft.app
    if hasattr(ft, "app"):
        ft.app(target=main)
    else:
        ft.run(target=main)
