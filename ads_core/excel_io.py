"""작업 엑셀 읽기(input 폴더) / 결과를 output 폴더의 사본 엑셀에 기록 / 현황 엑셀 만들기.

원본(input)은 읽기만 하고 고치지 않는다. 그래서 input 파일을 엑셀로 열어 둔 채로도 실행할 수 있다.
"""

import datetime as dt
import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from . import specs

RESULT_COLORS = {'완료': '188038', '예정': '5F6368'}   # 그 외(실패·오류)는 빨강


def _text(v):
    """셀 값 → 문자열. 엑셀이 날짜로 바꿔 둔 칸은 2026-10-01 형식으로, 300.0 같은 숫자는 300 으로 돌려준다."""
    if v is None:
        return ''
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_jobs(path):
    """작업 엑셀에서 템플릿 시트(키워드·소재·확장소재 6종)를 읽는다. 있는 시트만 돌려준다.
    반환: {시트 이름: {'headers': [1행 열 이름...], 'rows': [{'row': 행 번호, 'values': {열 이름: 값}}]}} (완전히 빈 행 제외)"""
    wb = load_workbook(path, data_only=True)
    out = {}
    try:
        for name in specs.BY_SHEET:
            if name not in wb.sheetnames:
                continue
            rows_iter = wb[name].iter_rows(values_only=True)
            header = next(rows_iter, None)
            if not header:
                continue
            headers = [_text(h) for h in header]
            rows = []
            for i, row in enumerate(rows_iter, start=2):
                values = {h: _text(v) for h, v in zip(headers, row) if h}
                if any(values.values()):
                    rows.append({'row': i, 'values': values})
            out[name] = {'headers': [h for h in headers if h], 'rows': rows}
    finally:
        wb.close()
    return out


class ResultWriter:
    """입력 엑셀의 사본에 결과·처리일시를 적어 output 폴더에 저장한다(서식·드롭다운은 그대로 유지)."""

    def __init__(self, src, dst):
        self.src, self.dst = src, dst
        self.wb = load_workbook(src)
        self._cols = {}

    def add(self, results):
        """results: [{'sheet': 시트 이름, 'row': 행 번호, 'result': 문구, 'time': 처리일시(선택)}]"""
        for r in results:
            if r['sheet'] not in self.wb.sheetnames:
                continue
            ws = self.wb[r['sheet']]
            cols = self._cols.setdefault(r['sheet'], {str(c.value).strip(): c.column for c in ws[1] if c.value})
            if '결과' not in cols:
                continue
            cell = ws.cell(r['row'], cols['결과'])
            cell.value = r['result']
            cell.font = Font(color=RESULT_COLORS.get(r['result'][:2], 'D93025'))
            if r.get('time') and '처리일시' in cols:
                ws.cell(r['row'], cols['처리일시']).value = r['time']

    def save(self):
        """저장하고 실제 경로를 돌려준다. 같은 이름의 결과 파일이 엑셀에서 열려 있어 저장이 막히면 이름 뒤에 번호를 붙인다."""
        os.makedirs(os.path.dirname(self.dst), exist_ok=True)
        stem, ext = os.path.splitext(self.dst)
        for n in range(0, 20):
            path = self.dst if n == 0 else f'{stem}_{n}{ext}'
            try:
                self.wb.save(path)
                return path
            except PermissionError:
                continue
        raise PermissionError(f'결과 파일을 저장하지 못했습니다(열려 있는 파일을 닫아 주세요): {self.dst}')


def write_snapshot_file(path, tables):
    """현황 엑셀을 만든다. tables: {'키워드': 2차원 배열(첫 행 머리글), ...} → 시트 이름은 대상 이름 그대로('키워드', '소재', …).
    파일 이름이 이미 "현황_…" 이고, 작업 엑셀의 시트 이름과 같아야 값을 옮기기 쉽다."""
    wb = Workbook()
    wb.remove(wb.active)
    for kind, rows in tables.items():
        ws = wb.create_sheet(kind)
        for row in rows:
            ws.append(row)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = 'A2'
        if len(rows) > 1:
            ws.auto_filter.ref = ws.dimensions
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path
