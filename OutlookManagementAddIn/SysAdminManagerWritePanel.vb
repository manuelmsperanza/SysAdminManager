Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook

Public Class SysAdminManagerWritePanel

    Private NoDeferredDelivery As New Date(4501, 1, 1) ' Magic number Outlook uses for "delay mail box isn't checked"
    Private NoTaskDates As New Date(1899, 12, 30) ' Magic number Outlook uses for "task dates"

#Region "Form Region Factory"

    <Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)>
    <Microsoft.Office.Tools.Outlook.FormRegionName("OutlookManagementAddIn.SysAdminManagerWritePanel")>
    Partial Public Class SysAdminManagerWritePanelFactory

        ' Occurs before the form region is initialized.
        ' To prevent the form region from appearing, set e.Cancel to true.
        ' Use e.OutlookItem to get a reference to the current Outlook item.
        Private Sub SysAdminManagerWritePanelFactory_FormRegionInitializing(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs) Handles Me.FormRegionInitializing

        End Sub

    End Class

#End Region

    'Occurs before the form region is displayed. 
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub SysAdminManagerWritePanel_FormRegionShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionShowing
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        Me.RetrieveFutureAppointments(mailItem)

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

    Private Sub RetrieveFutureAppointments(ByRef mailItem As MailItem)
        'mailItem.Body = "RetrieveFutureAppointments" & vbNewLine
        Dim sendDate As DateTime = Now
        'mailItem.Body += sendDate + vbNewLine
        Dim scheduleSendDate As Boolean = False
        If sendDate.Hour >= 18 Then
            sendDate = sendDate.AddDays(1).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
            'mailItem.Body += "after 18 " & sendDate & vbNewLine
            scheduleSendDate = True
        ElseIf sendDate.Hour < 9 Then
            sendDate = sendDate.AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
            'mailItem.Body += "before 9 " & sendDate & vbNewLine
            scheduleSendDate = True
        End If

        Select Case sendDate.DayOfWeek
            Case DayOfWeek.Saturday
                sendDate = sendDate.AddDays(2).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                'mailItem.Body += "saturday " & sendDate & vbNewLine
                scheduleSendDate = True
            Case DayOfWeek.Sunday
                sendDate = sendDate.AddDays(1).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                'mailItem.Body += "sunday " & sendDate & vbNewLine
                scheduleSendDate = True
        End Select

        Dim oCalendar As Outlook.Folder = mailItem.GetInspector.Application.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderCalendar)
        Dim oItems As Outlook.Items = oCalendar.Items

        ' Filter for appointments ending today or later
        Dim strFilter As String = "[End] >= '" & Today & "' and [End] < '" & Today.AddYears(1) & "' and [BusyStatus] = 'Out of Office'"
        oItems = oItems.Restrict(strFilter)
        oItems.IncludeRecurrences = True
        ' Sort the appointments
        oItems.Sort("[Start]")

        ' Loop through filtered items and print details
        For Each oAppointment As Outlook.AppointmentItem In oItems
            'mailItem.Body += "oAppointment " & oAppointment.Subject & " from " & oAppointment.Start & " till " & oAppointment.End & vbNewLine
            If sendDate.CompareTo(oAppointment.Start) >= 0 And sendDate.CompareTo(oAppointment.End) <= 0 Then
                sendDate = oAppointment.End
                'mailItem.Body += "ooo " & sendDate & vbNewLine
                scheduleSendDate = True

                If sendDate.Hour >= 18 Then
                    sendDate = sendDate.AddDays(1).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                    'mailItem.Body += "after 18 " & sendDate & vbNewLine
                    scheduleSendDate = True
                ElseIf sendDate.Hour < 9 Then
                    sendDate = sendDate.AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                    'mailItem.Body += "before 9 " & sendDate & vbNewLine
                    scheduleSendDate = True
                End If

                Select Case sendDate.DayOfWeek
                    Case DayOfWeek.Saturday
                        sendDate = sendDate.AddDays(2).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                        'mailItem.Body += "saturday " & sendDate & vbNewLine
                        scheduleSendDate = True
                    Case DayOfWeek.Sunday
                        sendDate = sendDate.AddDays(1).AddHours(9 - sendDate.Hour).AddMinutes(sendDate.Minute * -1)
                        scheduleSendDate = True
                        'mailItem.Body += "sunday " & sendDate & vbNewLine
                End Select

            Else
                Exit For
            End If

        Next oAppointment

        Me.DelayerDateTimePicker.Checked = scheduleSendDate
        If scheduleSendDate Then
            Me.DelayerDateTimePicker.Value = sendDate
            mailItem.DeferredDeliveryTime = sendDate
        Else
            mailItem.DeferredDeliveryTime = NoDeferredDelivery
        End If


    End Sub


    'Occurs when the form region is closed.   
    'Use Me.OutlookItem to get a reference to the current Outlook item.
    'Use Me.OutlookFormRegion to get a reference to the form region.
    Private Sub SysAdminManagerWritePanel_FormRegionClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FormRegionClosed

    End Sub

    Private Sub DelayerDateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles DelayerDateTimePicker.ValueChanged

        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        If Me.DelayerDateTimePicker.Checked Then
            mailItem.DeferredDeliveryTime = Me.DelayerDateTimePicker.Value
        Else
            mailItem.DeferredDeliveryTime = NoDeferredDelivery 'Date.MinValue
        End If

    End Sub

    Private Sub setCategories()
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.Categories = Me.OriginComboBox.Text & "," & Me.EffortTypeComboBox.Text
        mailItem.Save()
    End Sub

    Private Sub OriginComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OriginComboBox.SelectedIndexChanged
        Me.setCategories()
    End Sub

    Private Sub EffortType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EffortTypeComboBox.SelectedIndexChanged
        Me.setCategories()
    End Sub

    Private Sub SendImmediatlyButton_Click(sender As Object, e As EventArgs) Handles SendImmediatlyButton.Click
        Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(Me.OutlookItem, Microsoft.Office.Interop.Outlook.MailItem)
        mailItem.DeferredDeliveryTime = NoDeferredDelivery
        mailItem.Send()
    End Sub
End Class
