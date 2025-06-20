Attribute VB_Name = "Module26"
Sub DownloadUSPatentPDFs()
    Dim patentNo As String
    Dim formattedPatentNo As String
    Dim downloadURL As String
    Dim savePath As String
    Dim folderPath As String
    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim cell As Range
    Dim fd As FileDialog

    ' Ask user to select the folder once
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select a Folder to Save PDFs"
        If .Show <> -1 Then Exit Sub 'User cancelled
        folderPath = .SelectedItems(1)
    End With

    ' Loop through each selected cell
    For Each cell In Selection
        patentNo = Trim(cell.Value)

        If patentNo <> "" Then
            downloadURL = "https://image-ppubs.uspto.gov/dirsearch-public/print/downloadPdf/" & FormatPatentNumber(patentNo)
            savePath = folderPath & "\" & patentNo & ".pdf"

            ' Initialize HTTP Request
            Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
            WinHttpReq.Open "GET", downloadURL, False
            WinHttpReq.send

            If WinHttpReq.Status = 200 Then
                Set oStream = CreateObject("ADODB.Stream")
                oStream.Open
                oStream.Type = 1
                oStream.Write WinHttpReq.responseBody
                oStream.SaveToFile savePath, 2
                oStream.Close
                cell.Offset(0, 1).Value = "Downloaded"
            Else
                cell.Offset(0, 1).Value = "Failed: HTTP " & WinHttpReq.Status
            End If
        Else
            cell.Offset(0, 1).Value = "Empty Cell"
        End If
    Next cell

    MsgBox "PDF download process completed.", vbInformation

End Sub
Function FormatPatentNumber(patentNumber As String) As String
    Dim countryCode As String, kindCode As String, numberPart As String
    Dim i As Integer

    countryCode = Left(patentNumber, 2)
    numberPart = ""
    kindCode = ""

    For i = 3 To Len(patentNumber)
        If Mid(patentNumber, i, 1) Like "[A-Za-z]" Then
            kindCode = Mid(patentNumber, i)
            Exit For
        Else
            numberPart = numberPart & Mid(patentNumber, i, 1)
        End If
    Next i

    If kindCode = "" Then kindCode = "A1" ' default kind code if none provided

    FormatPatentNumber = numberPart
End Function

