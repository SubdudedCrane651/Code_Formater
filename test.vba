
Sub AddButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim btn As Button
    Dim i As Long
    For i = 2 To lastRow
        Set btn = ws.Buttons.Add(ws.Cells(i, 10).Left, ws.Cells(i, 9).Top, 100, 20)
        btn.Name = "btnShowImage" & i
        btn.OnAction = "ShowImage"
        ' Set the button caption to the value in Column A
        btn.Caption = ws.Cells(i, 1).Value
    Next i
End Sub

Sub ShowImage()
    ' In the References dialog, scroll down and check the box for
    ' Microsoft Visual Basic for Applications Extensibility 5.3
    ' & Microsoft Forms 2.0 Object Library by inserting a form
    Dim btnName As String
    btnName = Application.Caller
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim rowIndex As Long
    rowIndex = ws.Shapes(btnName).TopLeftCell.Row
    
    Dim imageURL As String
    imageURL = ws.Cells(rowIndex, 9).Value
    
    If imageURL <> "N/A" Then
        ' Download the image from the URL
        Dim localFilePath As String
        localFilePath = DownloadImage(imageURL)

        If localFilePath <> "" Then
            ' Create the UserForm dynamically
            Dim VBComp As VBComponent
            Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
            With VBComp
                .Properties("Width") = 320
                .Properties("Height") = 390
                .Properties("Caption") = ws.Cells(rowIndex, 2).Value
            End With

            ' Add an Image control to the UserForm
            Dim ImgControl As MSForms.Image
            Set ImgControl = VBComp.Designer.Controls.Add("Forms.Image.1")
            With ImgControl
                .Left = 10
                .Top = 10
                .Width = 320
                .Height = 350
                .Picture = LoadPicture(localFilePath)
            End With

            ' Show the dynamically created UserForm and delete it afterwards
            With VBA.UserForms.Add(VBComp.Name)
                .Show
                ' Once the form is closed, remove it from the workbook
                ThisWorkbook.VBProject.VBComponents.Remove VBComp
            End With
        Else
            MsgBox "Failed to download the image.", vbExclamation
        End If
    Else
        MsgBox "No image available for this movie.", vbExclamation
    End If
End Sub

Function DownloadImage(ByVal url As String) As String
    Dim httpReq As Object
    Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpReq.Open "GET", url, False
    httpReq.Send
    
    If httpReq.Status = 200 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1 ' Binary
        stream.Write httpReq.responseBody
        stream.SaveToFile Environ("TEMP") & "\downloaded_image.jpg", 2 ' Overwrite if exists
        stream.Close
        DownloadImage = Environ("TEMP") & "\downloaded_image.jpg"
    Else
        DownloadImage = ""
    End If
End Function
