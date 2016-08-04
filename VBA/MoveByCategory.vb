Sub MoveByCategory(Msg As Outlook.MailItem)

    Dim addIn As COMAddIn
    Dim automationObject As Object
    Set addIn = Application.COMAddIns("OutlookArchiveByCategoryAddIn")
    Set automationObject = addIn.Object
    If Not automationObject Is Nothing Then
        automationObject.ArchiveMailItem Msg
    End If
   

End Sub
