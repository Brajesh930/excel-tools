Attribute VB_Name = "Module19"
Sub match()
Dim mybro As Selenium.ChromeDriver
Set mybro = New Selenium.ChromeDriver
Dim webs As Selenium.WebElements
Dim web As Selenium.WebElement
Dim arrSplitStrings1 As Variant
Dim By As Selenium.By
Set By = New Selenium.By
'mybro.AddArgument ("headless")
mybro.start
mybro.Window.Maximize

link1 = "https://logicapt.keka.com/#/me/attendance/logs"
mybro.Get link1

mybro.Wait (10)
jump1:
XPath1 = "/html/body/xhr-app-root/div/employee-me/div/employee-attendance/div/div/div/div/employee-attendance-logs/div/employee-attendance-list-view/div/div[2]/div[1]/div/div[1]/div[2]/div/div[2]/div[2]/span/text()"
'On Error GoTo jump1
'Debug.Print mybro.FindElementByXPath(XPath1, timeout:=20000).Text
'On Error GoTo jump1
'ActiveCell.Value = mybro.FindElementByClass("score", timeout:=20000).Text
ActiveCell.Value = mybro.FindElementByXPath(XPath1, timeout:=20000).Text
mybro.Wait (2000)
GoTo jump1

End Sub
