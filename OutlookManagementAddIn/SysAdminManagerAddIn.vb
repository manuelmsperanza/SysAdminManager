Public Class SysAdminManagerAddIn
    Private WithEvents inspectors As Outlook.Inspectors
    Private olNs As Outlook.NameSpace
    Private NoDeferredDelivery As New Date(4501, 1, 1) ' Magic number Outlook uses for "delay mail box isn't checked"

    Private Sub SysAdminManagerAddIn_Startup() Handles Me.Startup
        Me.inspectors = Me.Application.Inspectors
        Me.olNs = Me.Application.GetNamespace("MAPI")
    End Sub

    Private Sub SysAdminManagerAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub inspectors_NewInspector(ByVal Inspector As Microsoft.Office.Interop.Outlook.Inspector) Handles inspectors.NewInspector
        Console.WriteLine("inspectors_NewInspector")
        Dim mailItem As Outlook.MailItem = TryCast(Inspector.CurrentItem, Outlook.MailItem)
        If Not (mailItem Is Nothing) Then
            If mailItem.EntryID Is Nothing Then
                'mailItem.Subject = "This text was added by using code"
                'mailItem.Body = "This text was added by using code"
            End If
        End If
    End Sub

    Private Sub Application_ItemSend(Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend

        Dim mailItem As Outlook.MailItem = TryCast(Item, Outlook.MailItem)
        If Not (mailItem Is Nothing) Then
            If mailItem.DeferredDeliveryTime.CompareTo(NoDeferredDelivery) < 0 And mailItem.DeferredDeliveryTime.CompareTo(Now) > 0 Then
                Dim msgBoxStyle As MsgBoxStyle = MsgBoxStyle.Question Or MsgBoxStyle.OkCancel Or MsgBoxStyle.DefaultButton1
                Dim msgBoxTitle As String = "Delayed Send"
                Dim result As MsgBoxResult = MsgBox("Item will be sent at : " & mailItem.DeferredDeliveryTime & ". Are you sure sure?", msgBoxStyle, msgBoxTitle)
                Cancel = result = MsgBoxResult.Cancel
            End If
        End If

    End Sub

End Class
