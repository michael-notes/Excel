# Excel
excel 使用技巧

# 如何快速将工作表批量转换成工作簿
https://zhuanlan.zhihu.com/p/43555350
## VBA代码

Sub WorkbookToSheet()
     Application.DisplayAlerts = False
      Application.ScreenUpdating = False
      For i = 1 To ThisWorkbook.Sheets.Count
           ThisWorkbook.Sheets(i).Copy
          ActiveWorkbook.SaveAs ThisWorkbook.Path & "/" & ThisWorkbook.Sheets(i).Name, xlWorkbookDefault
           ActiveWorkbook.Close True
      Next
      Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "处理完成。", , "提醒"
End Sub
