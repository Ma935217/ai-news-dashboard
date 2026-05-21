# -*- coding: utf-8 -*-
"""
グループホーム向け 半自動シフト管理Excelファイル生成スクリプト
出力ファイル: shift_management.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName

# ===== カラーパレット =====
HEADER_BG = "2E4057"
HEADER_FG = "FFFFFF"
ZEBRA_BG = "F8F9FA"

SHIFT_COLORS = {
    "早": {"bg": "D4EDDA", "fg": "1B5E20"},
    "日": {"bg": "CCE5FF", "fg": "0D47A1"},
    "遅": {"bg": "FFF3CD", "fg": "7D6608"},
    "夜": {"bg": "E8DAEF", "fg": "6C3483"},
    "明": {"bg": "FDEBD0", "fg": "A04000"},
    "休": {"bg": "F2F3F4", "fg": "7F8C8D"},
    "有": {"bg": "D5F5E3", "fg": "1A5276"},
    "P":  {"bg": "EBF5FB", "fg": "154360"},
}

FONT_NAME = "Meiryo UI"

thin = Side(border_style="thin", color="BDC3C7")
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)


def header_style(cell):
    cell.font = Font(name=FONT_NAME, size=11, bold=True, color=HEADER_FG)
    cell.fill = PatternFill("solid", fgColor=HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER_ALL


def body_style(cell, align="center"):
    cell.font = Font(name=FONT_NAME, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = BORDER_ALL


def apply_zebra(ws, start_row, end_row, start_col, end_col):
    """偶数行に淡色背景を適用"""
    for r in range(start_row, end_row + 1):
        if (r - start_row) % 2 == 1:
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                if cell.fill.fgColor.rgb in (None, "00000000", "FFFFFFFF"):
                    cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)


# ===========================================================
# 1. 職員マスタシート
# ===========================================================
def build_staff_master(wb):
    ws = wb.create_sheet("職員マスタ")
    ws.sheet_view.showGridLines = False

    headers = [
        "No.", "職員名", "フリガナ", "雇用形態", "外国人区分",
        "介護福祉士", "社会福祉士", "ケアマネ", "その他資格",
        "連勤上限(日)", "夜勤可否", "備考"
    ]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        header_style(cell)

    # サンプルデータ15名（日本人12名・外国人3名）
    staff = [
        (1,  "佐藤 健一", "サトウ ケンイチ",   "正職員", "日本人", "○", "",  "",  "",                   5, "可", "管理者"),
        (2,  "鈴木 美咲", "スズキ ミサキ",     "正職員", "日本人", "○", "",  "",  "",                   5, "可", ""),
        (3,  "高橋 大輔", "タカハシ ダイスケ", "正職員", "日本人", "○", "",  "○","",                   5, "可", ""),
        (4,  "田中 由美", "タナカ ユミ",       "正職員", "日本人", "",  "○","",  "初任者研修",         5, "可", ""),
        (5,  "伊藤 翔太", "イトウ ショウタ",   "正職員", "日本人", "○", "",  "",  "",                   5, "可", ""),
        (6,  "渡辺 真理", "ワタナベ マリ",     "パート", "日本人", "",  "",  "",  "実務者研修",         4, "否", "週3勤務"),
        (7,  "山本 直樹", "ヤマモト ナオキ",   "正職員", "日本人", "○", "",  "",  "",                   5, "可", ""),
        (8,  "中村 さくら","ナカムラ サクラ",  "パート", "日本人", "",  "",  "",  "初任者研修",         4, "否", "週4勤務"),
        (9,  "小林 健二", "コバヤシ ケンジ",   "正職員", "日本人", "",  "",  "○","",                   5, "可", ""),
        (10, "加藤 麻衣", "カトウ マイ",       "正職員", "日本人", "○", "",  "",  "",                   5, "可", ""),
        (11, "吉田 太郎", "ヨシダ タロウ",     "パート", "日本人", "",  "",  "",  "",                   3, "否", "週2勤務"),
        (12, "山田 花子", "ヤマダ ハナコ",     "正職員", "日本人", "○", "",  "",  "",                   5, "可", ""),
        (13, "ビカス シュレスタ",  "ビカス シュレスタ",  "正職員", "外国人", "",  "",  "",  "EPA介護福祉士候補", 5, "可", "ネパール出身"),
        (14, "サンディヤ タマン",  "サンディヤ タマン",  "正職員", "外国人", "○", "",  "",  "",                  5, "可", "ネパール出身"),
        (15, "プラディープ グルン","プラディープ グルン","正職員", "外国人", "",  "",  "",  "実務者研修",        5, "可", "ネパール出身"),
    ]

    for r, row in enumerate(staff, start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            align = "left" if c in (2, 3, 9, 12) else "center"
            body_style(cell, align=align)

    # 列幅
    widths = [5, 18, 22, 10, 11, 11, 11, 9, 18, 12, 10, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ドロップダウン
    dv_emp = DataValidation(type="list", formula1='"正職員,パート,その他"', allow_blank=True)
    dv_for = DataValidation(type="list", formula1='"日本人,外国人"', allow_blank=True)
    dv_night = DataValidation(type="list", formula1='"可,否"', allow_blank=True)
    ws.add_data_validation(dv_emp)
    ws.add_data_validation(dv_for)
    ws.add_data_validation(dv_night)
    dv_emp.add(f"D2:D100")
    dv_for.add(f"E2:E100")
    dv_night.add(f"K2:K100")

    ws.freeze_panes = "C2"
    apply_zebra(ws, 2, 16, 1, len(headers))
    ws.row_dimensions[1].height = 28


# ===========================================================
# 2. シフト設定シート
# ===========================================================
def build_shift_settings(wb):
    ws = wb.create_sheet("シフト設定")
    ws.sheet_view.showGridLines = False

    # 勤務区分テーブル
    ws["B2"] = "■ 勤務区分テーブル"
    ws["B2"].font = Font(name=FONT_NAME, size=12, bold=True, color=HEADER_BG)

    shift_headers = ["区分コード", "表示名", "開始時刻", "終了時刻", "勤務時間", "背景色コード", "文字色コード"]
    for c, h in enumerate(shift_headers, start=2):
        cell = ws.cell(row=3, column=c, value=h)
        header_style(cell)

    shift_data = [
        ("早", "早番",       "07:00", "16:00", 8.0, SHIFT_COLORS["早"]["bg"], SHIFT_COLORS["早"]["fg"]),
        ("日", "日勤",       "09:00", "18:00", 8.0, SHIFT_COLORS["日"]["bg"], SHIFT_COLORS["日"]["fg"]),
        ("遅", "遅番",       "11:00", "20:00", 8.0, SHIFT_COLORS["遅"]["bg"], SHIFT_COLORS["遅"]["fg"]),
        ("夜", "夜勤(宿直)", "17:00", "10:00", 16.0, SHIFT_COLORS["夜"]["bg"], SHIFT_COLORS["夜"]["fg"]),
        ("明", "明け番",     "00:00", "10:00", 0.0, SHIFT_COLORS["明"]["bg"], SHIFT_COLORS["明"]["fg"]),
        ("休", "公休",       "",      "",      0.0, SHIFT_COLORS["休"]["bg"], SHIFT_COLORS["休"]["fg"]),
        ("有", "有給",       "",      "",      8.0, SHIFT_COLORS["有"]["bg"], SHIFT_COLORS["有"]["fg"]),
        ("P", "パート勤務",  "09:00", "13:00", 4.0, SHIFT_COLORS["P"]["bg"],  SHIFT_COLORS["P"]["fg"]),
    ]
    for r, row in enumerate(shift_data, start=4):
        for c, val in enumerate(row, start=2):
            cell = ws.cell(row=r, column=c, value=val)
            body_style(cell)

    # 配置ルールテーブル
    ws["B14"] = "■ 配置ルールテーブル"
    ws["B14"].font = Font(name=FONT_NAME, size=12, bold=True, color=HEADER_BG)

    rule_headers = ["ルール項目", "設定値"]
    for c, h in enumerate(rule_headers, start=2):
        cell = ws.cell(row=15, column=c, value=h)
        header_style(cell)

    rules = [
        ("夜勤者数（1日あたり）",        2),
        ("夜勤者のうち外国人上限",       1),
        ("日中最低出勤者数",             3),
        ("連勤上限（デフォルト）",       5),
        ("月夜勤回数の均等化",           "する"),
        ("有資格者(介護福祉士)最低配置数", 1),
    ]
    for r, (k, v) in enumerate(rules, start=16):
        c1 = ws.cell(row=r, column=2, value=k)
        c2 = ws.cell(row=r, column=3, value=v)
        body_style(c1, align="left")
        body_style(c2)

    # 列幅
    widths = {2: 30, 3: 16, 4: 14, 5: 14, 6: 12, 7: 16, 8: 16}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[3].height = 26
    ws.row_dimensions[15].height = 26


# ===========================================================
# 3. 希望シフト入力シート
# ===========================================================
def build_request_sheet(wb):
    ws = wb.create_sheet("希望シフト入力")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "対象年月"
    ws["A1"].font = Font(name=FONT_NAME, size=11, bold=True, color=HEADER_FG)
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = BORDER_ALL

    ws["B1"] = "2026/05"
    body_style(ws["B1"], align="left")
    ws["B1"].font = Font(name=FONT_NAME, size=11, bold=True)

    ws["D1"] = "※ 各セルにドロップダウンから希望区分を選択してください（早/日/遅/夜/明/休/有/P）"
    ws["D1"].font = Font(name=FONT_NAME, size=9, color="7F8C8D")
    ws["D1"].alignment = Alignment(horizontal="left", vertical="center")

    # ヘッダー（2行目）
    ws.cell(row=2, column=1, value="No.")
    ws.cell(row=2, column=2, value="職員名")
    header_style(ws.cell(row=2, column=1))
    header_style(ws.cell(row=2, column=2))
    for d in range(1, 32):
        cell = ws.cell(row=2, column=2 + d, value=d)
        header_style(cell)

    # 職員行（マスタから自動参照）
    for i in range(15):
        r = 3 + i
        ws.cell(row=r, column=1, value=f"=職員マスタ!A{2+i}")
        ws.cell(row=r, column=2, value=f"=職員マスタ!B{2+i}")
        body_style(ws.cell(row=r, column=1))
        body_style(ws.cell(row=r, column=2), align="left")
        for d in range(1, 32):
            cell = ws.cell(row=r, column=2 + d, value="")
            body_style(cell)

    # ドロップダウン
    dv = DataValidation(type="list", formula1='"早,日,遅,夜,明,休,有,P"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C3:AG17")

    # 列幅
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 18
    for d in range(1, 32):
        ws.column_dimensions[get_column_letter(2 + d)].width = 4.5

    ws.freeze_panes = "C3"
    ws.row_dimensions[2].height = 24
    apply_zebra(ws, 3, 17, 1, 33)


# ===========================================================
# 4. シフト表シート（メイン）
# ===========================================================
def build_shift_table(wb):
    ws = wb.create_sheet("シフト表")
    ws.sheet_view.showGridLines = False

    # タイトル
    ws["A1"] = "対象年月"
    ws["A1"].font = Font(name=FONT_NAME, size=11, bold=True, color=HEADER_FG)
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = BORDER_ALL

    ws["B1"] = "=希望シフト入力!B1"
    body_style(ws["B1"], align="left")
    ws["B1"].font = Font(name=FONT_NAME, size=11, bold=True)

    # ヘッダー（2行目）
    ws.cell(row=2, column=1, value="No.")
    ws.cell(row=2, column=2, value="職員名")
    header_style(ws.cell(row=2, column=1))
    header_style(ws.cell(row=2, column=2))
    for d in range(1, 32):
        cell = ws.cell(row=2, column=2 + d, value=d)
        header_style(cell)

    # 集計列ヘッダー（列34〜37）
    summary_cols = ["出勤日数", "夜勤回数", "公休日数", "有給日数"]
    for i, label in enumerate(summary_cols):
        cell = ws.cell(row=2, column=34 + i, value=label)
        header_style(cell)

    # 職員データ行（15名）
    for i in range(15):
        r = 3 + i
        ws.cell(row=r, column=1, value=f"=職員マスタ!A{2+i}")
        ws.cell(row=r, column=2, value=f"=職員マスタ!B{2+i}")
        body_style(ws.cell(row=r, column=1))
        body_style(ws.cell(row=r, column=2), align="left")

        for d in range(1, 32):
            cell = ws.cell(row=r, column=2 + d, value="")
            body_style(cell)

        # 集計列 (COUNTIF)
        rng = f"C{r}:AG{r}"
        # 出勤日数：早/日/遅/夜/明/P
        ws.cell(row=r, column=34,
                value=f'=COUNTIF({rng},"早")+COUNTIF({rng},"日")+COUNTIF({rng},"遅")+COUNTIF({rng},"夜")+COUNTIF({rng},"明")+COUNTIF({rng},"P")')
        ws.cell(row=r, column=35, value=f'=COUNTIF({rng},"夜")')
        ws.cell(row=r, column=36, value=f'=COUNTIF({rng},"休")')
        ws.cell(row=r, column=37, value=f'=COUNTIF({rng},"有")')
        for c in range(34, 38):
            body_style(ws.cell(row=r, column=c))
            ws.cell(row=r, column=c).font = Font(name=FONT_NAME, size=10, bold=True)

    # 下部集計行
    last_staff_row = 17  # 3〜17
    label_row_defs = [
        ("夜勤人数", "夜"),
        ("早番人数", "早"),
        ("日勤人数", "日"),
        ("遅番人数", "遅"),
    ]
    base = 18
    for idx, (label, code) in enumerate(label_row_defs):
        r = base + idx
        c1 = ws.cell(row=r, column=2, value=label)
        body_style(c1, align="left")
        c1.font = Font(name=FONT_NAME, size=10, bold=True)
        c1.fill = PatternFill("solid", fgColor="ECF0F1")
        for d in range(1, 32):
            col = 2 + d
            colL = get_column_letter(col)
            cell = ws.cell(row=r, column=col,
                           value=f'=COUNTIF({colL}3:{colL}{last_staff_row},"{code}")')
            body_style(cell)

    # 出勤合計行
    r_total = base + 4  # 22
    c_total = ws.cell(row=r_total, column=2, value="出勤合計")
    body_style(c_total, align="left")
    c_total.font = Font(name=FONT_NAME, size=10, bold=True)
    c_total.fill = PatternFill("solid", fgColor="ECF0F1")
    for d in range(1, 32):
        col = 2 + d
        colL = get_column_letter(col)
        formula = (
            f'=COUNTIF({colL}3:{colL}{last_staff_row},"早")'
            f'+COUNTIF({colL}3:{colL}{last_staff_row},"日")'
            f'+COUNTIF({colL}3:{colL}{last_staff_row},"遅")'
            f'+COUNTIF({colL}3:{colL}{last_staff_row},"夜")'
            f'+COUNTIF({colL}3:{colL}{last_staff_row},"明")'
            f'+COUNTIF({colL}3:{colL}{last_staff_row},"P")'
        )
        cell = ws.cell(row=r_total, column=col, value=formula)
        body_style(cell)
        cell.font = Font(name=FONT_NAME, size=10, bold=True)

    # 夜勤人数確認行（2名未満で「不足」）
    r_check = r_total + 1  # 23
    c_check = ws.cell(row=r_check, column=2, value="夜勤確認")
    body_style(c_check, align="left")
    c_check.font = Font(name=FONT_NAME, size=10, bold=True, color="C0392B")
    c_check.fill = PatternFill("solid", fgColor="FADBD8")
    for d in range(1, 32):
        col = 2 + d
        colL = get_column_letter(col)
        cell = ws.cell(row=r_check, column=col,
                       value=f'=IF({colL}18<2,"不足","")')
        body_style(cell)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color="C0392B")

    # ドロップダウン
    dv = DataValidation(type="list", formula1='"早,日,遅,夜,明,休,有,P"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C3:AG17")

    # 列幅
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 18
    for d in range(1, 32):
        ws.column_dimensions[get_column_letter(2 + d)].width = 4.5
    for c in range(34, 38):
        ws.column_dimensions[get_column_letter(c)].width = 11

    ws.freeze_panes = "C3"
    ws.row_dimensions[2].height = 24

    # 条件付き書式：シフト種別ごとに色付け
    shift_range = "C3:AG17"
    for code, colors in SHIFT_COLORS.items():
        rule = CellIsRule(
            operator="equal",
            formula=[f'"{code}"'],
            fill=PatternFill("solid", fgColor=colors["bg"]),
            font=Font(name=FONT_NAME, size=10, bold=True, color=colors["fg"]),
        )
        ws.conditional_formatting.add(shift_range, rule)

    # 「不足」セルを赤強調
    rule_short = CellIsRule(
        operator="equal",
        formula=['"不足"'],
        fill=PatternFill("solid", fgColor="E74C3C"),
        font=Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF"),
    )
    ws.conditional_formatting.add(f"C{r_check}:AG{r_check}", rule_short)


# ===========================================================
# 5. 月次サマリーシート
# ===========================================================
def build_monthly_summary(wb):
    ws = wb.create_sheet("月次サマリー")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "対象年月"
    ws["A1"].font = Font(name=FONT_NAME, size=11, bold=True, color=HEADER_FG)
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].border = BORDER_ALL

    ws["B1"] = "=シフト表!B1"
    body_style(ws["B1"], align="left")
    ws["B1"].font = Font(name=FONT_NAME, size=11, bold=True)

    headers = ["No.", "職員名", "出勤日数", "夜勤回数", "早番", "日勤", "遅番", "公休", "有給"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        header_style(cell)

    # 職員ごとの集計
    for i in range(15):
        r = 3 + i
        shift_row = 3 + i
        ws.cell(row=r, column=1, value=f"=職員マスタ!A{2+i}")
        ws.cell(row=r, column=2, value=f"=職員マスタ!B{2+i}")
        rng = f"シフト表!C{shift_row}:AG{shift_row}"
        ws.cell(row=r, column=3,
                value=f'=COUNTIF({rng},"早")+COUNTIF({rng},"日")+COUNTIF({rng},"遅")+COUNTIF({rng},"夜")+COUNTIF({rng},"明")+COUNTIF({rng},"P")')
        ws.cell(row=r, column=4, value=f'=COUNTIF({rng},"夜")')
        ws.cell(row=r, column=5, value=f'=COUNTIF({rng},"早")')
        ws.cell(row=r, column=6, value=f'=COUNTIF({rng},"日")')
        ws.cell(row=r, column=7, value=f'=COUNTIF({rng},"遅")')
        ws.cell(row=r, column=8, value=f'=COUNTIF({rng},"休")')
        ws.cell(row=r, column=9, value=f'=COUNTIF({rng},"有")')

        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            align = "left" if c == 2 else "center"
            body_style(cell, align=align)

    # 合計行
    r_total = 18
    ws.cell(row=r_total, column=1, value="")
    ws.cell(row=r_total, column=2, value="合計")
    for c in range(3, 10):
        colL = get_column_letter(c)
        ws.cell(row=r_total, column=c, value=f'=SUM({colL}3:{colL}17)')
    for c in range(1, 10):
        cell = ws.cell(row=r_total, column=c)
        align = "left" if c == 2 else "center"
        body_style(cell, align=align)
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="ECF0F1")

    widths = [5, 18, 11, 11, 9, 9, 9, 9, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "C3"
    ws.row_dimensions[2].height = 26
    apply_zebra(ws, 3, 17, 1, 9)


# ===========================================================
# 6. マクロコードシート
# ===========================================================
VBA_CODE = '''Option Explicit

'==========================================================
' グループホーム シフト管理マクロ
' 使用前に .xlsm 形式で保存し、本コードをVBAエディタに貼り付け
'==========================================================

' --- 設定定数 ---
Const SHEET_MASTER As String = "職員マスタ"
Const SHEET_SETTINGS As String = "シフト設定"
Const SHEET_REQUEST As String = "希望シフト入力"
Const SHEET_SHIFT As String = "シフト表"

Const STAFF_FIRST_ROW As Long = 3
Const STAFF_LAST_ROW As Long = 17
Const DAY_FIRST_COL As Long = 3      ' C列
Const DAY_LAST_COL As Long = 33      ' AG列
Const MAX_MONTH_DAYS As Long = 20    ' 月最大出勤日数

'----------------------------------------------------------
' メイン処理：希望シフトをシフト表へ転記し、空欄を自動割当
'----------------------------------------------------------
Public Sub CreateShift()
    Dim wsReq As Worksheet, wsShift As Worksheet, wsMaster As Worksheet
    Set wsReq = ThisWorkbook.Worksheets(SHEET_REQUEST)
    Set wsShift = ThisWorkbook.Worksheets(SHEET_SHIFT)
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_MASTER)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ① 希望シフトを転記
    Dim r As Long, c As Long
    For r = STAFF_FIRST_ROW To STAFF_LAST_ROW
        For c = DAY_FIRST_COL To DAY_LAST_COL
            wsShift.Cells(r, c).Value = wsReq.Cells(r, c).Value
        Next c
    Next r

    ' ② 各職員の属性を取得
    Dim staffCount As Long
    staffCount = STAFF_LAST_ROW - STAFF_FIRST_ROW + 1

    ReDim isPart(1 To staffCount) As Boolean
    ReDim canNight(1 To staffCount) As Boolean
    ReDim isForeign(1 To staffCount) As Boolean
    ReDim conLimit(1 To staffCount) As Long

    Dim i As Long
    For i = 1 To staffCount
        Dim mr As Long: mr = 1 + i           ' 職員マスタの行 (2-16)
        isPart(i) = (wsMaster.Cells(mr, 4).Value = "パート")
        canNight(i) = (wsMaster.Cells(mr, 11).Value = "可")
        isForeign(i) = (wsMaster.Cells(mr, 5).Value = "外国人")
        If IsNumeric(wsMaster.Cells(mr, 10).Value) Then
            conLimit(i) = CLng(wsMaster.Cells(mr, 10).Value)
        Else
            conLimit(i) = 5
        End If
    Next i

    ' ③ 自動割当：日ごとに夜勤2名を確保→空欄を埋める
    Dim d As Long
    For d = DAY_FIRST_COL To DAY_LAST_COL
        Call AssignNightShift(wsShift, d, isPart, canNight, isForeign)
    Next d

    ' ④ 夜勤翌日を「明」、連勤上限超過は「休」、月20日超は「休」
    Call FillRemaining(wsShift, isPart, canNight, conLimit)

    ' ⑤ 色分け
    Call ReColorize

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' ⑥ 夜勤不足日チェック
    Dim shortDays As Long: shortDays = 0
    For d = DAY_FIRST_COL To DAY_LAST_COL
        If CountShiftInColumn(wsShift, d, "夜") < 2 Then
            shortDays = shortDays + 1
        End If
    Next d

    If shortDays > 0 Then
        MsgBox "シフト作成完了。夜勤不足日: " & shortDays & "日。手動で調整してください。", _
               vbExclamation, "シフト作成"
    Else
        MsgBox "シフト作成完了。夜勤不足はありません。", vbInformation, "シフト作成"
    End If
End Sub

'----------------------------------------------------------
' 1日分の夜勤を2名（外国人は最大1名）まで割当
'----------------------------------------------------------
Private Sub AssignNightShift(ws As Worksheet, dayCol As Long, _
                              isPart() As Boolean, canNight() As Boolean, _
                              isForeign() As Boolean)
    Dim nightCnt As Long, foreignNight As Long
    Dim i As Long, r As Long

    ' 既存夜勤者を数える
    For i = 1 To UBound(canNight)
        r = STAFF_FIRST_ROW + i - 1
        If ws.Cells(r, dayCol).Value = "夜" Then
            nightCnt = nightCnt + 1
            If isForeign(i) Then foreignNight = foreignNight + 1
        End If
    Next i

    ' 不足分を埋める
    For i = 1 To UBound(canNight)
        If nightCnt >= 2 Then Exit For
        r = STAFF_FIRST_ROW + i - 1
        If ws.Cells(r, dayCol).Value = "" Then
            If isPart(i) Then GoTo NextI
            If Not canNight(i) Then GoTo NextI
            If isForeign(i) And foreignNight >= 1 Then GoTo NextI
            ' 前日が夜勤なら割当不可
            If dayCol > DAY_FIRST_COL Then
                If ws.Cells(r, dayCol - 1).Value = "夜" Then GoTo NextI
            End If
            ws.Cells(r, dayCol).Value = "夜"
            nightCnt = nightCnt + 1
            If isForeign(i) Then foreignNight = foreignNight + 1
        End If
NextI:
    Next i
End Sub

'----------------------------------------------------------
' 残りの空欄を埋める（明・休のルール処理含む）
'----------------------------------------------------------
Private Sub FillRemaining(ws As Worksheet, _
                           isPart() As Boolean, canNight() As Boolean, _
                           conLimit() As Long)
    Dim i As Long, d As Long, r As Long
    For i = 1 To UBound(canNight)
        r = STAFF_FIRST_ROW + i - 1
        Dim consec As Long: consec = 0
        Dim monthlyWork As Long: monthlyWork = 0

        For d = DAY_FIRST_COL To DAY_LAST_COL
            Dim prev As String: prev = ""
            If d > DAY_FIRST_COL Then prev = CStr(ws.Cells(r, d - 1).Value)

            ' 夜勤翌日は強制的に「明」
            If prev = "夜" Then
                If ws.Cells(r, d).Value = "" Or ws.Cells(r, d).Value = "明" Then
                    ws.Cells(r, d).Value = "明"
                End If
            End If

            If ws.Cells(r, d).Value = "" Then
                ' パートはP、夜勤不可職員はP/日勤
                If isPart(i) Then
                    ws.Cells(r, d).Value = "P"
                ElseIf monthlyWork >= MAX_MONTH_DAYS Then
                    ws.Cells(r, d).Value = "休"
                ElseIf consec >= conLimit(i) Then
                    ws.Cells(r, d).Value = "休"
                Else
                    ' 早/日/遅をローテーション
                    Select Case (d Mod 3)
                        Case 0: ws.Cells(r, d).Value = "早"
                        Case 1: ws.Cells(r, d).Value = "日"
                        Case 2: ws.Cells(r, d).Value = "遅"
                    End Select
                End If
            End If

            ' 連勤・月勤務カウント更新
            Dim v As String: v = CStr(ws.Cells(r, d).Value)
            If v = "休" Or v = "有" Or v = "" Then
                consec = 0
            Else
                consec = consec + 1
                monthlyWork = monthlyWork + 1
            End If
        Next d
    Next i
End Sub

'----------------------------------------------------------
' 列内のシフトコード件数
'----------------------------------------------------------
Private Function CountShiftInColumn(ws As Worksheet, col As Long, code As String) As Long
    Dim cnt As Long: cnt = 0
    Dim r As Long
    For r = STAFF_FIRST_ROW To STAFF_LAST_ROW
        If ws.Cells(r, col).Value = code Then cnt = cnt + 1
    Next r
    CountShiftInColumn = cnt
End Function

'----------------------------------------------------------
' シフトクリア
'----------------------------------------------------------
Public Sub ClearShift()
    If MsgBox("シフト表をすべてクリアします。よろしいですか？", _
              vbYesNo + vbQuestion, "確認") <> vbYes Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SHIFT)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(STAFF_FIRST_ROW, DAY_FIRST_COL), _
                       ws.Cells(STAFF_LAST_ROW, DAY_LAST_COL))
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    rng.Font.Color = RGB(0, 0, 0)

    MsgBox "シフト表をクリアしました。", vbInformation
End Sub

'----------------------------------------------------------
' 色を再適用（条件付き書式と二重掛けの場合のフェイルセーフ）
'----------------------------------------------------------
Public Sub ReColorize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SHIFT)

    Dim r As Long, c As Long, v As String
    For r = STAFF_FIRST_ROW To STAFF_LAST_ROW
        For c = DAY_FIRST_COL To DAY_LAST_COL
            v = CStr(ws.Cells(r, c).Value)
            Select Case v
                Case "早": Call ApplyColor(ws.Cells(r, c), "D4EDDA", "1B5E20")
                Case "日": Call ApplyColor(ws.Cells(r, c), "CCE5FF", "0D47A1")
                Case "遅": Call ApplyColor(ws.Cells(r, c), "FFF3CD", "7D6608")
                Case "夜": Call ApplyColor(ws.Cells(r, c), "E8DAEF", "6C3483")
                Case "明": Call ApplyColor(ws.Cells(r, c), "FDEBD0", "A04000")
                Case "休": Call ApplyColor(ws.Cells(r, c), "F2F3F4", "7F8C8D")
                Case "有": Call ApplyColor(ws.Cells(r, c), "D5F5E3", "1A5276")
                Case "P":  Call ApplyColor(ws.Cells(r, c), "EBF5FB", "154360")
                Case Else
                    ws.Cells(r, c).Interior.ColorIndex = xlNone
            End Select
        Next c
    Next r
End Sub

Private Sub ApplyColor(cell As Range, bg As String, fg As String)
    cell.Interior.Color = HexToLong(bg)
    cell.Font.Color = HexToLong(fg)
    cell.Font.Bold = True
End Sub

Private Function HexToLong(hex As String) As Long
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(hex, 1, 2))
    g = CLng("&H" & Mid(hex, 3, 2))
    b = CLng("&H" & Mid(hex, 5, 2))
    HexToLong = RGB(r, g, b)
End Function
'==========================================================
' END OF MACRO
'=========================================================='''


def build_macro_sheet(wb):
    ws = wb.create_sheet("マクロコード(参考)")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "VBAマクロコード（コピー用）"
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    # 使用方法
    ws["A3"] = "■ 使用方法"
    ws["A3"].font = Font(name=FONT_NAME, size=12, bold=True, color=HEADER_BG)

    usage = [
        "1. このファイルを .xlsm 形式（マクロ有効ブック）で保存し直してください。",
        "   [ファイル] → [名前を付けて保存] → ファイル形式 \"Excel マクロ有効ブック (*.xlsm)\"",
        "2. Excelで Alt + F11 キーを押し、VBAエディタを開きます。",
        "3. メニュー [挿入] → [標準モジュール] を選択し、新規モジュールを作成。",
        "4. 下記のVBAコード全文をモジュールに貼り付けて保存。",
        "5. シフト表シートに戻り、Alt + F8 でマクロ一覧を開き、",
        "   ・CreateShift … 希望シフトを基にシフト表を自動生成",
        "   ・ClearShift  … シフト表をクリア（確認ダイアログあり）",
        "   ・ReColorize  … 手動修正後の色を再適用",
        "   を実行してください。",
    ]
    for i, line in enumerate(usage, start=4):
        ws.cell(row=i, column=1, value=line)
        ws.cell(row=i, column=1).font = Font(name=FONT_NAME, size=10)
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # コード本体
    ws["A16"] = "■ VBAコード（以下を全てコピーしてください）"
    ws["A16"].font = Font(name=FONT_NAME, size=12, bold=True, color=HEADER_BG)

    lines = VBA_CODE.split("\n")
    for i, line in enumerate(lines, start=17):
        cell = ws.cell(row=i, column=1, value=line)
        cell.font = Font(name="Consolas", size=10)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, indent=1)
        # コメント行は色付け
        if line.strip().startswith("'") or line.strip().startswith("Const") or line.strip().startswith("'==="):
            cell.font = Font(name="Consolas", size=10, color="7F8C8D")
        if line.strip().startswith("Public Sub") or line.strip().startswith("Private Sub") or line.strip().startswith("Private Function"):
            cell.font = Font(name="Consolas", size=10, bold=True, color="1F618D")

    ws.column_dimensions["A"].width = 110


# ===========================================================
# メイン
# ===========================================================
def main():
    wb = Workbook()
    # デフォルトシートを削除
    default = wb.active
    wb.remove(default)

    build_staff_master(wb)
    build_shift_settings(wb)
    build_request_sheet(wb)
    build_shift_table(wb)
    build_monthly_summary(wb)
    build_macro_sheet(wb)

    out = "shift_management.xlsx"
    wb.save(out)
    print(f"✅ 生成完了: {out}")


if __name__ == "__main__":
    main()
