Attribute VB_Name = "ETL"
Option Explicit

Const Path As String = "c:\users\futur\Desktop\abc\"
Dim AttName As String, Name1 As String, Name2 As String, Name3 As String, Name4 As String, Name5 As String, Name6 As String

Sub Extract()

Dim Ol As Outlook.Application, ONS As Outlook.Namespace, Folder As Outlook.MAPIFolder, Item As Object, Mail As Outlook.Mailitem, Att As Outlook.Attachment, Filter As String, Start As Double

Start = [now()]

With Application
    .ScreenUpdating = False
    .Visible = False
End With

    Filter = "[ReceivedTime]>'" & Format(Date, "DDDD HHH:NN") & "'"
    AttName = "VBA-test-2.xlsm"
    
    Set Ol = New Outlook.Application
    Set ONS = Ol.GetNamespace("MAPI")
    Set Folder = ONS.Folders("krzysztof.gajewski.1@tlen.pl")
    Set Folder = Folder.Folders("Inbox")
    
        For Each Item In Folder.Items.Restrict(Filter)
            If Item.Class = OlMail Then
                Set Mail = Item
                    If Mail.Attachments.Count > 0 Then
                        For Each Att In Item.Attachments
                            If InStr(Att.Filename, AttName) Then
                                Att.SaveAsFile Path & Att.Filename
                                Call Transform
                            End If
                         Next Att
                   End If
             End If
        Next Item

With Application
    .ScreenUpdating = True
    .Visible = True
End With

MsgBox Application.WorksheetFunction.Text([now()] - Start, "h:mm:ss.00")

End Sub

Sub Transform()

Dim wb As Workbook, ws As Worksheet, i As Integer, exists As Boolean, LR As Integer, j As Integer

    Set wb = Workbooks.Open(Path & AttName)

    For i = 1 To wb.Worksheets.Count
        If Worksheets(i).name = "vba2 raw Name4" Then
            exists = True
            Exit For
        End If
    Next i
    
    If exists = False Then
        With Application
            .ScreenUpdating = True
            .Visible = True
        End With
            MsgBox "The specific worksheet not found. Please double check the source file"
            Exit Sub
    Else
        Set ws = wb.Worksheets("vba2 raw Name4")
    End If
    
    If ws.Range("A1") <> "Unique Code" Or InStr(1, ws.Range("A2"), "5521") = 0 Then
        With Application
            .ScreenUpdating = True
            .Visible = True
        End With
            MsgBox "Error. Please double check the source file"
            Exit Sub
    End If
    
    LR = ws.Range("B" & Rows.Count).End(xlUp).Row
    
        For i = 2 To LR
                
                If InStr(1, ws.Range("B" & i), "Unique Code ") > 0 Then
                    Name1 = ws.Range("B" & i).Value
                        Name1 = Replace(Name1, "Unique Code ", vbNullString)
                            Name1 = SoapSyntax(Name1)
                End If
                
                If IsNumeric(ws.Range("BB" & i)) = True And ws.Range("BB" & i) <> vbNullString Then
                    Name2 = ws.Range("BB" & i)
                        Name2 = SoapSyntax(Name2)
                    Name3 = ws.Range("AB" & i)
                        Name3 = SoapSyntax(Name3)
                    Name4 = ws.Range("Q" & i)
                        Name4 = SoapSyntax(Name4)
                    Name5 = ws.Range("BN" & i)
                        Name5 = SoapSyntax(Name5)
                    Name6 = ws.Range("AQ" & i)
                        Name6 = SoapSyntax(Name6)
                End If
                
                Call Load
                
        Next i

End Sub


Sub Load()

Dim ListName As String ' sharepoint listname
ListName = "presentation"
Dim SharepointUrl As String: SharepointUrl = "https://sharepointaddress/sites/name%20name%20/" 'sharepoint list address

Dim objXMLHTTP As Object 'the xml object with the post request
Dim strBatchXml As String 'part of the body of the post request (with all the list fields and values)
Dim strSoapBody As String 'full post requet with all the tags

    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    
    strBatchXml = "<Batch onError='Continue'><Method ID='3' Cmd='New'>" & _
    "<Field Name='Unique Code'>" & Name1 & "</Field>" & _
    "<Field Name='Name'>" & Name2 & "</Field>" & _
    "<Field Name='Comments'>" & Name3 & "</Field>" & _
    "<Field Name='Value'>" & Name4 & "</Field>" & _
    "<Field Name='Data'>" & Name5 & "</Field>" & _
    "<Field Name='Updated_x002f_Created_x0020_by'>" & Name6 & "</Field>" & _
    "</Method></Batch"
    
    objXMLHTTP.Open "POST", SharepointUrl + "_vti_bin/Lists.asmx", False
    objXMLHTTP.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
    objXMLHTTP.setRequestHeader "SOAPaction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
    
    'insert the body of the post request
    strSoapBody = "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " _
    & "xmlns:xsd='htttp://www.w3.org/2001/XMLSchema' " _
    & "xmlsn:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><UpdateListItems" _
    & "xmlns='http://schemas.microsoft.com/sharepoint/soap/'><ListName>" & ListName _
    & "</listName><updates>" & strBatchXml & "</updates></UpdateListItems></soap:Body></soap:Envelope>"
    
    objXMLHTTP.Send strSoapBody ' send request
    Set objXMLHTTP = Nothing ' clear object variable
    
End Sub

Function SoapSyntax(name As String)

    name = Replace(name, "&", "&amp;")
    name = Replace(name, Chr(39), "&apos;")
    name = Replace(name, "<", "&lt;")
    name = Replace(name, ">", "&gt;")
    name = Replace(name, Chr(34), "&quot;")
    
    SoapSyntax = name
    
End Function
