# 패키지 불러오기
from __future__ import annotations
import win32com.client



# 엑셀 애플리케이션 준비
def initExcel() -> win32com.client.CDispatch:
    excel = win32com.client.Dispatch("Excel.Application")
    return excel