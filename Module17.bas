Attribute VB_Name = "Module17"
Sub Legalstatus()
Dim mybro As Selenium.ChromeDriver
Set mybro = New Selenium.ChromeDriver
Dim webs As Selenium.WebElements
Dim web As Selenium.WebElement
Dim By As Selenium.By
Dim pr As Single
Set By = New Selenium.By
mybro.AddArgument ("headless")
mybro.start
mybro.Window.Maximize
ufProgress.LabelProgress.Width = 0
ufProgress.Show
ncell = Selection.Cells.count
pr = 0

For Each cell In Selection

FractionComplete (pr)
If IsEmpty(ActiveCell) = True Then
Dim answer As Integer
answer = MsgBox("Cell is empty in  " & ActiveCell.row & "  row and " & ActiveCell.Column & "  column", vbQuestion + vbYesNo + vbDefaultButton2, "Do you want to exit")
    If answer = vbYes Then
    Unload ufProgress
    Exit Sub
     Else: GoTo jump2
     End If
    End If

link1 = "https://patents.google.com/patent/" + CStr(ActiveCell.Value)
mybro.Get (link1)
On Error GoTo ErrorHandler:
link3 = mybro.FindElementByPartialLinkText("Privacy Policy", timeout:=20000).Attribute("href")
On Error GoTo ErrorHandler:
link2 = Replace(mybro.FindElementByPartialLinkText("USPTO PatentCenter", timeout:=1000).Attribute("href"), "#!/", "")
ActiveCell.Offset(0, 1).Select
mybro.Get (link2)
On Error GoTo ErrorHandler1:
ActiveCell.Value = mybro.FindElementByXPath("/html/body/div[3]/main/div/div/div/div/div[2]/div[2]/div[4]/div/span[1]", timeout:=20000).Text
mybro.Wait (500)
'End If
jump1:

ActiveCell.Offset(0, -1).Select
jump2:
ActiveCell.Offset(1, 0).Select
pr = pr + (1 / ncell)
Next cell
ErrorHandler:
If Err.Number <> 0 Then
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Error _ Non US patent or slow internet"
Err.Clear
Resume jump1
End If
ErrorHandler1:
If Err.Number <> 0 Then
'ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Error _ Non US patent or slow internet"
Err.Clear
Resume jump1
End If
Unload ufProgress
End Sub
Sub FractionComplete(pctdone As Single)
With ufProgress
    .LabelCaption.Caption = pctdone * 100 & "% Complete"
    .LabelProgress.Width = pctdone * (.FrameProgress.Width)
End With
DoEvents
End Sub

