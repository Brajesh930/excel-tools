Attribute VB_Name = "Module15"
Sub forwardcitations()
Dim mybro As Selenium.ChromeDriver
Set mybro = New Selenium.ChromeDriver
Dim webs As Selenium.WebElements
Dim web As Selenium.WebElement
Dim By As Selenium.By
Set By = New Selenium.By
Dim arrSplitStrings1 As Variant
mybro.start
Dim rng As Range, cell As Range
Set rng = Selection
For Each cell In rng
If IsEmpty(ActiveCell) = True Then
Dim answer As Integer
answer = MsgBox("Cell is empty in  " & ActiveCell.row & "  row and " & ActiveCell.Column & "  column", vbQuestion + vbYesNo + vbDefaultButton2, "Do you want to exit")
    If answer = vbYes Then
     Exit For
     Else: GoTo jump
     End If
    End If
link1 = "https://worldwide.espacenet.com/patent/search?q=ct%3d" + ActiveCell.Text
mybro.Get link1
mybro.Window.Maximize
a = MsgBox("wait for page load and then click ok", vbOKOnly, "Your internet speed")
mybro.Wait (10)
ActiveCell.Offset(0, 32).Select
xpath2 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div[1]/div[4]/article[1]/section/div[1]/span[1]"
If Not mybro.IsElementPresent(By.XPath(xpath2)) Then
xpath2 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[2]/section[1]/span/a/span"
ActiveCell.Value = mybro.FindElementByXPath(xpath2).Text
GoTo jump1
End If
For i = 1 To 700
XPath1 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div[1]/div[4]/article[" + CStr(i) + "]/section/header/div"
xpath2 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div[1]/div[4]/article[" + CStr(i) + "]/section/div[1]/span[1]"
If Not mybro.IsElementPresent(By.XPath(XPath1)) Then
If i < 2 Then
ActiveCell.Value = "NA"
GoTo jump1
Else: GoTo jump1
End If
End If
arrSplitStrings1 = Split(CStr(mybro.FindElementByXPath(xpath2).Text), " ")
On Error GoTo jump1
ActiveCell.Value = ActiveCell.Value & arrSplitStrings1(0) & Chr(10)
'Debug.Print arrSplitStrings1(0)
mybro.FindElementByXPath(XPath1).Click
On Error GoTo jump1
mybro.Wait (40)
Next
jump1:
mybro.Wait (400)
jump:
ActiveCell.Offset(1, -32).Select
Next cell
End Sub
