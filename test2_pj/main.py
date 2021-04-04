import win32com.client
from pywintypes import com_error
from pathlib import Path


# 検索する文字列: 置換する文字列
REPLACE_TXTS = {
    '{{セルテキスト1}}': 'こんにちは。',
    '{{セルテキスト2}}': 'おはようございます。'
}

nowdir = Path(__file__).absolute().parent

excel = win32com.client.Dispatch('Excel.Application')

excel.Visible = False

try:
    # xlsxファイルを開く
    wb = excel.Workbooks.Add(str(nowdir / 'template.xlsx'))
    sheet = wb.WorkSheets(1)
    sheet.Activate()

    # セル内のテキストを置換
    rg = sheet.Range(sheet.usedRange.Address)
    for search_txt, replace_Txt in REPLACE_TXTS.items():
        rg.Replace(search_txt, replace_Txt)

    # 別名で保存
    wb.SaveAs(str(nowdir / 'new.xlsx'))

except com_error as e:
    print('失敗。', e)
else:
    print('成功。')
finally:
    wb.Close(False)
    excel.Quit()