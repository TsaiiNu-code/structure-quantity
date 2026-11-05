"""Create an explicitly uncalculated workbook; no project quantities are assumed."""
import argparse
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def create_template(output, floors):
    wb = Workbook()
    wb.remove(wb.active)

    def sheet(name, headers, rows=()):
        # Excel names are limited to 31 characters and must be unique ignoring case.
        base = re.sub(r'[\\/*?:\[\]]', '_', name).strip("'")[:31] or '工作表'
        title = base
        count = 1
        while title.casefold() in {s.casefold() for s in wb.sheetnames}:
            suffix = f'_{count}'
            title = base[:31-len(suffix)] + suffix
            count += 1
        ws = wb.create_sheet(title)
        ws.append(headers)
        for row in rows:
            ws.append(row)
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='234E70')
            cell.font = Font(color='FFFFFF', bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[1].height = 32
        for i, header in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = min(44, max(18, len(header)*2+2))
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = '1:1'
        return ws

    sheet('說明', ['項目', '內容'], [
        ['狀態', '待填範本；尚未取得來源檔，無實際工程數量。空白不代表零。'],
        ['構件範圍', '柱、梁、板、RC牆；依種類與實際樓層分表。'],
        ['RC牆依據', '指定 DXF 的 RC 牆圖層；實際名稱與 IFC 對應待確認。'],
        ['混凝土', '已確認以體積 m³ 計量。'],
        ['模板', '實際接觸面 m²；交接面、開口及施工界面扣除規則待確認。'],
        ['鋼筋', '依圖說計算長度與核定單位重；搭接、錨定、彎鉤、損耗分列。'],
        ['樓層', '、'.join(floors) if floors else '未提供；構件分表名稱「待定樓層」只是佔位。'],
        ['追溯', '保留來源檔、版本、圖號、GUID 與計算依據。'],
    ])
    sheet('來源清單', ['來源ID', '檔案類型', '檔名或位置', '版本或雜湊', '圖號', '適用樓層', '單位', '狀態'])
    sheet('RC牆圖層', ['來源ID', 'DXF圖層原名', '樓層', '圖元Handle', 'IFC GUID', '對應方法', '納入RC牆', '覆核狀態', '備註'])
    sheet('樓層彙總', ['樓層', '種類', '模板面積(m²)', '混凝土體積(m³)', '來源分表', '覆核狀態'])
    sheet('鋼筋彙總', ['樓層', '種類', '鋼筋號數', '總長(m)', '重量(kg)', '損耗另計(kg)', '單位重依據', '覆核狀態'])
    for floor in floors or ['待定樓層']:
        for kind in ['柱', '梁', '板', 'RC牆']:
            sheet(f'{floor}_{kind}', [
                '樓層', '種類', '構件編號', 'IFC GUID', '來源ID及圖號', '版本',
                'RC牆圖層對應', '長(m)', '寬或厚(m)', '高(m)', '構件數',
                '模板計算式', '模板毛面積(m²)', '模板扣除(m²)', '模板淨面積(m²)',
                '混凝土體積計算式',
                '混凝土毛體積(m³)', '混凝土扣除(m³)', '混凝土淨體積(m³)', '覆核狀態', '備註'])
        sheet(f'{floor}_鋼筋', [
            '樓層', '種類', '構件編號', 'IFC GUID', '來源ID及圖號', '版本', '筋別',
            '鋼筋號數', '間距(mm)', '支數計算式', '支數', '單支基本長(m)',
            '單支搭接長(m)', '單支錨定長(m)', '單支彎鉤長(m)', '單支下料長(m)',
            '總長(m)', '單位重(kg/m)', '單位重依據', '重量(kg)', '損耗率(%)',
            '損耗另計(kg)', '續接器數', '覆核狀態', '備註'])
    sheet('待確認', ['問題ID', '問題', '影響範圍', '來源ID及圖號', '處理狀態', '核定內容及依據'], [
        ['Q001', '取得 IFC 與完整圖說', '全部數量', '', '待提供'],
        ['Q002', '混凝土以體積 m³ 計量', '混凝土', '', '使用者已確認'],
        ['Q003', '確認 RC 牆圖層名稱及模型對應', 'RC牆', '', '待取得DXF'],
        ['Q004', '確認樓層、開口及構件交疊扣除規則', '模板與混凝土', '', '待確認'],
        ['Q005', '確認鋼筋單位重、搭接錨定彎鉤及損耗規則', '鋼筋', '', '待取得圖說'],
    ])
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', default='outputs/結構數量_待填範本.xlsx')
    parser.add_argument('--floors', nargs='+', default=[])
    args = parser.parse_args()
    print(create_template(args.output, args.floors))
