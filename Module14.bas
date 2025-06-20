Attribute VB_Name = "Module14"
Sub fullreport()
Dim mybro As Selenium.ChromeDriver
Set mybro = New Selenium.ChromeDriver
Dim webs As Selenium.WebElements
Dim web As Selenium.WebElement
Dim arrSplitStrings1 As Variant
Dim priorities As Variant
Dim priority1 As Variant
Dim By As Selenium.By
Set By = New Selenium.By
ReDim arrSplitStrings1(0 To 4)
Dim myDate As Date
Dim mydate1 As Date
Dim publicationdate As Date
Dim applicationdate As Date
myDate = Date
mybro.start
mybro.Window.Maximize
Dim rng As Range, cell As Range
Set rng = Selection
'link1 = mybro.baseUrl

For Each cell In rng
If IsEmpty(ActiveCell) = True Then
Dim answer As Integer
answer = MsgBox("Cell is empty in  " & ActiveCell.row & "  row and " & ActiveCell.Column & "  column", vbQuestion + vbYesNo + vbDefaultButton2, "Do you want to exit")
    If answer = vbYes Then
     Exit For
     Else: GoTo jump1
     End If
    End If
link1 = "https://worldwide.espacenet.com/patent/search/" + ActiveCell.Text + "?q=" + ActiveCell.Text
myDate = Date
patentNo = ActiveCell.Text
'Debug.Print patentno
patentNo = Trim(patentNo)
'Debug.Print patentno
cucode = Left(ActiveCell.Text, 2)
mybro.Get link1
a = MsgBox("wait for page load and then click ok", vbOKOnly, "Your internet speed")
ActiveCell.Offset(0, 3).Select
ActiveCell.Value = mybro.FindElementById("biblio-title-content").Text
On Error Resume Next

For j = 0 To 50
If Not mybro.IsElementPresent(By.ID("biblio-priority-numbers-content-" + CStr(j))) Then Exit For
prioritie = mybro.FindElementById("biblio-priority-numbers-content-" + CStr(j)).Text
prioritynumbers = prioritynumbers + Left(prioritie, (InStr(1, prioritie, "·") - 1)) + ";"
mydate1 = CDate(Right(prioritie, (Len(prioritie) - (InStr(1, prioritie, "·")))))
If myDate > mydate1 Then
myDate = mydate1
prioritycountry = Left(prioritie, 2)
End If
Next

Application1 = mybro.FindElementById("biblio-application-number-content").Text
applicationnumber = Left(Application1, (InStr(1, Application1, "·") - 1))
applicationdate = CDate(Right(Application1, (Len(Application1) - (InStr(1, Application1, "·")))))

Publication = mybro.FindElementById("biblio-publication-number-content").Text
Publicationnumber = Left(Publication, (InStr(1, Publication, "·") - 1))
publicationdate = CDate(Right(Publication, (Len(Publication) - (InStr(1, Publication, "·")))))

ActiveCell.Offset(0, 1).Select
ActiveCell.Value = mybro.FindElementById("biblio-abstract-content").Text
mybro.Wait (100)
On Error Resume Next
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = mybro.FindElementById("biblio-applicants-content").Text
On Error Resume Next
inventorname = mybro.FindElementById("biblio-inventors-content").Text
On Error Resume Next
cpc = mybro.FindElementById("biblio-cooperative-content").Text
On Error Resume Next
cpc = Replace(cpc, "CPC" + Chr(10), "")
'Debug.Print patentno
relatedapplications = mybro.FindElementById("biblio-also-published-as-content").Text
relatedapplication = Replace(relatedapplications, patentNo + ";", "")
relatedapplication = Replace(relatedapplications, patentNo, "")

'Debug.Print patentno
On Error Resume Next
mybro.FindElementByXPath("/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[2]/ul/li[8]/span").Click
a = MsgBox("wait for page load,go to simplefamily, and then click ok", vbOKOnly, "Your internet speed")
mybro.Wait (10)
ActiveCell.Offset(0, 3).Select

For i = 1 To 700
XPath1 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[3]/div[1]/div[3]/div/div[2]/table/tbody/tr[" + CStr(i) + "]/td[1]/a/span"
If Not mybro.IsElementPresent(By.XPath(XPath1)) Then GoTo jumpa1
'Debug.Print mybro.FindElementByXPath(xpath1).Text
On Error GoTo jumpa1
If Len(mybro.FindElementByXPath(XPath1).Text) < 3 Then GoTo jump3
arrSplitStrings1 = Split(CStr(mybro.FindElementByXPath(XPath1).Text), " ")
On Error GoTo jump3
'Debug.Print arrSplitStrings1(0)
ActiveCell.Value = ActiveCell.Value & arrSplitStrings1(0) & Chr(10)
jump3:
mybro.Wait (40)
Next

jumpa1:
mybro.FindElementByXPath("/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[3]/div[1]/div[2]/ul/li[2]/span").Click
a = MsgBox("wait for page load,go to inpadocfamily, and then click ok", vbOKOnly, "Your internet speed")
mybro.Wait (10)
ActiveCell.Offset(0, 1).Select

For i = 1 To 700
XPath1 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[3]/div[1]/div[3]/div[2]/div[2]/table/tbody/tr[" + CStr(i) + "]/td[1]/a/span"
If Not mybro.IsElementPresent(By.XPath(XPath1)) Then GoTo jumpa2
'Debug.Print mybro.FindElementByXPath(xpath1).Text
On Error GoTo jumpa2
If Len(mybro.FindElementByXPath(XPath1).Text) < 3 Then GoTo jump23
arrSplitStrings1 = Split(CStr(mybro.FindElementByXPath(XPath1).Text), " ")
On Error GoTo jump23
'Debug.Print arrSplitStrings1(0)
ActiveCell.Value = ActiveCell.Value & arrSplitStrings1(0) & Chr(10)
jump23:
mybro.Wait (40)
Next

jumpa2:
ActiveCell.Offset(0, 2).Select
ActiveCell.Value = cucode


ActiveCell.Offset(0, 2).Select
ActiveCell.Value = inventorname


ActiveCell.Offset(0, 1).Select
If Not (publicationdate = Date) Then
ActiveCell.Value = Year(publicationdate)
End If

ActiveCell.Offset(0, 1).Select
If Not (publicationdate = Date) Then
ActiveCell.Value = publicationdate
End If


ActiveCell.Offset(0, 1).Select
If Not (applicationdate = Date) Then
ActiveCell.Value = applicationdate
End If


ActiveCell.Offset(0, 1).Select
ActiveCell.Value = applicationnumber


ActiveCell.Offset(0, 1).Select
ActiveCell.Value = relatedapplication


ActiveCell.Offset(0, 1).Select
ActiveCell.Value = myDate

ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Year(myDate)

ActiveCell.Offset(0, 1).Select
ActiveCell.Value = prioritycountry

ActiveCell.Offset(0, 1).Select
ActiveCell.Value = prioritynumbers
ActiveCell.Offset(0, 8).Select
ActiveCell.Value = cpc

mybro.FindElementByXPath("/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[2]/ul/li[6]/span").Click
a = MsgBox("wait for page load,go to citation, and then click ok", vbOKOnly, "Your internet speed")
mybro.Wait (10)
ActiveCell.Offset(0, 1).Select

For i = 1 To 700
XPath1 = "/html/body/div[1]/div/div[2]/div/div[3]/div/div/div[3]/div[3]/div/div[2]/div/div[2]/table/tbody/tr[" + CStr(i) + "]/td[2]/span/span"
If Not mybro.IsElementPresent(By.XPath(XPath1)) Then GoTo jump1
'Debug.Print mybro.FindElementByXPath(xpath1).Text
On Error GoTo jump1
If Len(mybro.FindElementByXPath(XPath1).Text) < 3 Then GoTo jumpb3
arrSplitStrings1 = Split(CStr(mybro.FindElementByXPath(XPath1).Text), " ")
On Error GoTo jumpb3
'Debug.Print arrSplitStrings1(0)
ActiveCell.Value = ActiveCell.Value & arrSplitStrings1(0) & Chr(10)
jumpb3:
mybro.Wait (40)
Next

mybro.Wait (400)
jump1:
cpc = ""
prioritynumbers = ""
prioritycountry = ""
relatedapplication = ""
applicationnumber = ""
applicationdate = Date
publicationdate = Date
inventorname = ""
cucode = ""
ActiveCell.Offset(1, -31).Select

Next cell

End Sub
