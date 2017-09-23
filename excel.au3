#include <Date.au3>
#include <Excel.au3>


; connect to excel and open new instance
Local $oExcel= _Excel_Open()

;Local $sWorkbook = "C:\THE\PATH\LOCATION\Book1.xlsx"

; create new workbook
Local $oWorkbook = _Excel_BookNew($oExcel)

; fill in data
_Excel_RangeWrite($oWorkbook, Default, "**", "B3:D10")

With $oWorkbook.ActiveSheet
    Local $iFirstRow = .usedRange.Row
    Local $iLastRow = $iFirstRow + .usedRange.rows.count -1
    for $iRow = $iFirstRow to $iLastRow Step 2
        .rows($iRow).interior.colorIndex = 36
    next
endWith

sleep(1000)

; change focus to target window
WinActivate("windowtitle")

msgbox(0, "the script", "all done at: " & _NowTime())

; Aut2exe.exe /in C:\Users\HOME\Desktop\dev\bots\excelMacro.au3 /out C:\Users\HOME\Desktop\dev\bots\excelMacro.exe /x64
