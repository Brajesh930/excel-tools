Attribute VB_Name = "Module13"
Sub inpadocfamily()
Dim mybro As Selenium.ChromeDriver
Set mybro = New Selenium.ChromeDriver
Dim webs As Selenium.WebElements
Dim web As Selenium.WebElement
Dim arrSplitStrings1 As Variant
Dim By As Selenium.By
Set By = New Selenium.By
ReDim arrSplitStrings1(0 To 4)
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
link1 = "https://worldwide.espacenet.com/patent/search/" + ActiveCell.Text + "?q=" + ActiveCell.Text
mybro.Get link1
a = MsgBox("wait for page load,go to inpadocfamily, and then click ok", vbOKOnly, "Your internet speed")
mybro.Wait (10)
ActiveCell.Offset(0, 3).Select
For i = 1 To 700
XPath1 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[3]/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[" + CStr(i) + "]/td[1]/a/span"
If Not mybro.IsElementPresent(By.XPath(XPath1)) Then GoTo jump1
Debug.Print mybro.FindElementByXPath(XPath1).Text
On Error GoTo jump1
If Len(mybro.FindElementByXPath(XPath1).Text) < 3 Then GoTo jump3
arrSplitStrings1 = Split(CStr(mybro.FindElementByXPath(XPath1).Text), " ")
On Error GoTo jump3
Debug.Print arrSplitStrings1(0)
ActiveCell.Value = ActiveCell.Value & arrSplitStrings1(0) & Chr(10)
jump3:
mybro.Wait (40)
Next
jump1:
mybro.Wait (400)
jump:
ActiveCell.Offset(1, -3).Select
Next cell
End Sub

