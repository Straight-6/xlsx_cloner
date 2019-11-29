' .xlsx Cloner


Set App = CreateObject("Excel.Application")
App.Visible = False

Set objDoc = App.Workbooks.Open(WScript.Arguments(0))
Set objDoc_temp = App.Workbooks.Add

For i = 1 To objDoc.Worksheets.Count
    ' MsgBox(objDoc.Worksheets(i).Name)
    Set cnt = objDoc_temp.Worksheets(objDoc_temp.Worksheets.Count)
    Call objDoc.Worksheets(i).Copy(,cnt)
Next

For i = 1 To objDoc_temp.Worksheets.Count - objDoc.Worksheets.Count
    ' MsgBox(objDoc.Worksheets(i).Name)
    Call objDoc_temp.Worksheets(i).Delete
Next

fpath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%") & "/../Local/Temp/temp.xlsx"
objDoc_temp.SaveAs(fpath)

objDoc.Close
objDoc_temp.Close
App.Quit
