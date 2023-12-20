Imports Microsoft.Office.Interop.Outlook

Public Class SysAdminManagerReadPanel
    Private NoDeferredDelivery As New Date(4501, 1, 1) ' Magic number Outlook uses for "delay mail box isn't checked"
    Private NoTaskDates As New Date(1899, 12, 30) ' Magic number Outlook uses for "task dates"
#Region "Form Region Factory"

    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)>
    <Microsoft.Office.Tools.Outlook.FormRegionName("OutlookManagementAddIn.SysAdminManagerReadPanel")>
    Partial Public Class SysAdminManagerReadPanelFactory

        ' Occurs before the form region is initialized.
        ' To prevent the form region from appearing, set e.Cancel to true.
        ' Use e.OutlookItem to get a reference to the current Outlook item.
        Private Sub SysAdminManagerReadPanelFactory_FormRegionInitializing(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs) Handles Me.FormRegionInitializing

        End Sub

    End Class

#End Region

    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub SysAdminManagerReadPanel_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)

        If mailItem.Categories IsNot Nothing Then

            For Each curCategory As String In mailItem.Categories.Split(";")

                Dim idxOrigin As Integer = Me.OriginComboBox.Items.IndexOf(curCategory.Trim())
                If idxOrigin >= 0 Then
                    Me.OriginComboBox.SelectedItem = Me.OriginComboBox.Items.Item(idxOrigin)
                End If

                Dim idxEffortType As Integer = Me.EffortTypeComboBox.Items.IndexOf(curCategory.Trim())
                If idxEffortType >= 0 Then
                    Me.EffortTypeComboBox.SelectedItem = Me.EffortTypeComboBox.Items.Item(idxEffortType)
                End If

            Next curCategory

        End If

    End Sub

    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub SysAdminManagerReadPanel_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub

    Private Sub setCategories()
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.Categories = Me.OriginComboBox.Text & "," & Me.EffortTypeComboBox.Text

        If Me.EffortTypeComboBox.Text IsNot Nothing And Not mailItem.IsMarkedAsTask Then
            mailItem.MarkAsTask(Microsoft.Office.Interop.Outlook.OlMarkInterval.olMarkNoDate)
            mailItem.FlagRequest = "New"
        End If

        mailItem.Save()
    End Sub

    Private Sub OriginComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OriginComboBox.SelectedIndexChanged
        Me.setCategories()
    End Sub

    Private Sub EffortType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EffortTypeComboBox.SelectedIndexChanged
        Me.setCategories()
    End Sub

    Private Sub setFlagRequest(ByRef mailItem As Microsoft.Office.Interop.Outlook.MailItem, ByRef flagRequest As String)
        If mailItem.IsMarkedAsTask Then
            mailItem.MarkAsTask(Microsoft.Office.Interop.Outlook.OlMarkInterval.olMarkNoDate)
        End If
        mailItem.FlagRequest = flagRequest
    End Sub

    Private Sub NewTaskButton_Click(sender As Object, e As EventArgs) Handles NewTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        Me.setFlagRequest(mailItem, "New")
        mailItem.Save()
    End Sub

    Private Sub BacklogTaskButton_Click(sender As Object, e As EventArgs) Handles BacklogTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        Me.setFlagRequest(mailItem, "Backlog")
        mailItem.Save()
    End Sub

    Private Sub ActiveTaskButton_Click(sender As Object, e As EventArgs) Handles ActiveTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        Me.setFlagRequest(mailItem, "Active")
        mailItem.Save()
    End Sub

    Private Sub VerifyingTaskButton_Click(sender As Object, e As EventArgs) Handles VerifyingTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        Me.setFlagRequest(mailItem, "Verifying")
        mailItem.Save()
    End Sub

    Private Sub CompletedTaskButton_Click(sender As Object, e As EventArgs) Handles CompletedTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.MarkAsTask(OlMarkInterval.olMarkComplete)
        mailItem.Save()
    End Sub

    Private Sub NoDateTaskButton_Click(sender As Object, e As EventArgs) Handles NoDateTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.MarkAsTask(OlMarkInterval.olMarkNoDate)
        mailItem.Save()
    End Sub

    Private Sub TodayTaskButton_Click(sender As Object, e As EventArgs) Handles TodayTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.MarkAsTask(OlMarkInterval.olMarkToday)
        mailItem.Save()
    End Sub

    Private Sub ThisWeekButton_Click(sender As Object, e As EventArgs) Handles ThisWeekButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.MarkAsTask(OlMarkInterval.olMarkThisWeek)
        mailItem.Save()
    End Sub

    Private Sub NextWeekTaskButton_Click(sender As Object, e As EventArgs) Handles NextWeekTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.MarkAsTask(OlMarkInterval.olMarkNextWeek)
        mailItem.Save()
    End Sub

    Private Sub ResetTaskButton_Click(sender As Object, e As EventArgs) Handles ResetTaskButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.ClearTaskFlag()
        mailItem.Save()
    End Sub

End Class
