<System.ComponentModel.ToolboxItemAttribute(False)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SysAdminManagerReadPanel
    Inherits Microsoft.Office.Tools.Outlook.FormRegionBase

    Public Sub New(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion)
        MyBase.New(Globals.Factory, formRegion)
        Me.InitializeComponent()
    End Sub

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Form Regions Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Shared Sub InitializeManifest(ByVal manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest, ByVal factory As Microsoft.Office.Tools.Outlook.Factory)
        manifest.FormRegionName = "SysAdminManagerReadPanel"
        manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining
        manifest.ShowInspectorCompose = False

    End Sub

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.EffortTypeLabel = New System.Windows.Forms.Label()
        Me.OriginLabel = New System.Windows.Forms.Label()
        Me.EffortTypeComboBox = New System.Windows.Forms.ComboBox()
        Me.OriginComboBox = New System.Windows.Forms.ComboBox()
        Me.NoDateTaskButton = New System.Windows.Forms.Button()
        Me.NewTaskButton = New System.Windows.Forms.Button()
        Me.BacklogTaskButton = New System.Windows.Forms.Button()
        Me.ActiveTaskButton = New System.Windows.Forms.Button()
        Me.VerifyingTaskButton = New System.Windows.Forms.Button()
        Me.CompletedTaskButton = New System.Windows.Forms.Button()
        Me.TodayTaskButton = New System.Windows.Forms.Button()
        Me.ThisWeekButton = New System.Windows.Forms.Button()
        Me.NextWeekTaskButton = New System.Windows.Forms.Button()
        Me.ResetTaskButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'EffortTypeLabel
        '
        Me.EffortTypeLabel.AutoSize = True
        Me.EffortTypeLabel.Location = New System.Drawing.Point(6, 57)
        Me.EffortTypeLabel.Name = "EffortTypeLabel"
        Me.EffortTypeLabel.Size = New System.Drawing.Size(59, 13)
        Me.EffortTypeLabel.TabIndex = 15
        Me.EffortTypeLabel.Text = "Effort Type"
        '
        'OriginLabel
        '
        Me.OriginLabel.AutoSize = True
        Me.OriginLabel.Location = New System.Drawing.Point(31, 18)
        Me.OriginLabel.Name = "OriginLabel"
        Me.OriginLabel.Size = New System.Drawing.Size(34, 13)
        Me.OriginLabel.TabIndex = 14
        Me.OriginLabel.Text = "Origin"
        '
        'EffortTypeComboBox
        '
        Me.EffortTypeComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.EffortTypeComboBox.FormattingEnabled = True
        Me.EffortTypeComboBox.Items.AddRange(New Object() {"Quick Request", "Long Request", "Emergency"})
        Me.EffortTypeComboBox.Location = New System.Drawing.Point(71, 53)
        Me.EffortTypeComboBox.Name = "EffortTypeComboBox"
        Me.EffortTypeComboBox.Size = New System.Drawing.Size(200, 21)
        Me.EffortTypeComboBox.TabIndex = 13
        '
        'OriginComboBox
        '
        Me.OriginComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.OriginComboBox.FormattingEnabled = True
        Me.OriginComboBox.Items.AddRange(New Object() {"Project Task", "Emergency Request", "Incident", "Incident - alert", "Incident - outage"})
        Me.OriginComboBox.Location = New System.Drawing.Point(71, 14)
        Me.OriginComboBox.Name = "OriginComboBox"
        Me.OriginComboBox.Size = New System.Drawing.Size(200, 21)
        Me.OriginComboBox.TabIndex = 12
        '
        'NoDateTaskButton
        '
        Me.NoDateTaskButton.Location = New System.Drawing.Point(296, 52)
        Me.NoDateTaskButton.Name = "NoDateTaskButton"
        Me.NoDateTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.NoDateTaskButton.TabIndex = 16
        Me.NoDateTaskButton.Text = "No Date"
        Me.NoDateTaskButton.UseVisualStyleBackColor = True
        '
        'NewTaskButton
        '
        Me.NewTaskButton.Location = New System.Drawing.Point(296, 13)
        Me.NewTaskButton.Name = "NewTaskButton"
        Me.NewTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.NewTaskButton.TabIndex = 17
        Me.NewTaskButton.Text = "New"
        Me.NewTaskButton.UseVisualStyleBackColor = True
        '
        'BacklogTaskButton
        '
        Me.BacklogTaskButton.Location = New System.Drawing.Point(377, 13)
        Me.BacklogTaskButton.Name = "BacklogTaskButton"
        Me.BacklogTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.BacklogTaskButton.TabIndex = 18
        Me.BacklogTaskButton.Text = "Backlog"
        Me.BacklogTaskButton.UseVisualStyleBackColor = True
        '
        'ActiveTaskButton
        '
        Me.ActiveTaskButton.Location = New System.Drawing.Point(458, 13)
        Me.ActiveTaskButton.Name = "ActiveTaskButton"
        Me.ActiveTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.ActiveTaskButton.TabIndex = 19
        Me.ActiveTaskButton.Text = "Active"
        Me.ActiveTaskButton.UseVisualStyleBackColor = True
        '
        'VerifyingTaskButton
        '
        Me.VerifyingTaskButton.Location = New System.Drawing.Point(539, 13)
        Me.VerifyingTaskButton.Name = "VerifyingTaskButton"
        Me.VerifyingTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.VerifyingTaskButton.TabIndex = 20
        Me.VerifyingTaskButton.Text = "Verifying"
        Me.VerifyingTaskButton.UseVisualStyleBackColor = True
        '
        'CompletedTaskButton
        '
        Me.CompletedTaskButton.Location = New System.Drawing.Point(620, 13)
        Me.CompletedTaskButton.Name = "CompletedTaskButton"
        Me.CompletedTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.CompletedTaskButton.TabIndex = 21
        Me.CompletedTaskButton.Text = "Completed"
        Me.CompletedTaskButton.UseVisualStyleBackColor = True
        '
        'TodayTaskButton
        '
        Me.TodayTaskButton.Location = New System.Drawing.Point(377, 52)
        Me.TodayTaskButton.Name = "TodayTaskButton"
        Me.TodayTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.TodayTaskButton.TabIndex = 22
        Me.TodayTaskButton.Text = "Today"
        Me.TodayTaskButton.UseVisualStyleBackColor = True
        '
        'ThisWeekButton
        '
        Me.ThisWeekButton.Location = New System.Drawing.Point(458, 52)
        Me.ThisWeekButton.Name = "ThisWeekButton"
        Me.ThisWeekButton.Size = New System.Drawing.Size(75, 23)
        Me.ThisWeekButton.TabIndex = 23
        Me.ThisWeekButton.Text = "This Week"
        Me.ThisWeekButton.UseVisualStyleBackColor = True
        '
        'NextWeekTaskButton
        '
        Me.NextWeekTaskButton.Location = New System.Drawing.Point(539, 52)
        Me.NextWeekTaskButton.Name = "NextWeekTaskButton"
        Me.NextWeekTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.NextWeekTaskButton.TabIndex = 24
        Me.NextWeekTaskButton.Text = "Next Week"
        Me.NextWeekTaskButton.UseVisualStyleBackColor = True
        '
        'ResetTaskButton
        '
        Me.ResetTaskButton.Location = New System.Drawing.Point(620, 52)
        Me.ResetTaskButton.Name = "ResetTaskButton"
        Me.ResetTaskButton.Size = New System.Drawing.Size(75, 23)
        Me.ResetTaskButton.TabIndex = 25
        Me.ResetTaskButton.Text = "Reset"
        Me.ResetTaskButton.UseVisualStyleBackColor = True
        '
        'SysAdminManagerReadPanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ResetTaskButton)
        Me.Controls.Add(Me.NextWeekTaskButton)
        Me.Controls.Add(Me.ThisWeekButton)
        Me.Controls.Add(Me.TodayTaskButton)
        Me.Controls.Add(Me.CompletedTaskButton)
        Me.Controls.Add(Me.VerifyingTaskButton)
        Me.Controls.Add(Me.ActiveTaskButton)
        Me.Controls.Add(Me.BacklogTaskButton)
        Me.Controls.Add(Me.NewTaskButton)
        Me.Controls.Add(Me.NoDateTaskButton)
        Me.Controls.Add(Me.EffortTypeLabel)
        Me.Controls.Add(Me.OriginLabel)
        Me.Controls.Add(Me.EffortTypeComboBox)
        Me.Controls.Add(Me.OriginComboBox)
        Me.Name = "SysAdminManagerReadPanel"
        Me.Size = New System.Drawing.Size(714, 92)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents EffortTypeLabel As Windows.Forms.Label
    Friend WithEvents OriginLabel As Windows.Forms.Label
    Friend WithEvents EffortTypeComboBox As Windows.Forms.ComboBox
    Friend WithEvents OriginComboBox As Windows.Forms.ComboBox
    Friend WithEvents NoDateTaskButton As Windows.Forms.Button
    Friend WithEvents NewTaskButton As Windows.Forms.Button
    Friend WithEvents BacklogTaskButton As Windows.Forms.Button
    Friend WithEvents ActiveTaskButton As Windows.Forms.Button
    Friend WithEvents VerifyingTaskButton As Windows.Forms.Button
    Friend WithEvents CompletedTaskButton As Windows.Forms.Button
    Friend WithEvents TodayTaskButton As Windows.Forms.Button
    Friend WithEvents ThisWeekButton As Windows.Forms.Button
    Friend WithEvents NextWeekTaskButton As Windows.Forms.Button
    Friend WithEvents ResetTaskButton As Windows.Forms.Button

    Partial Public Class SysAdminManagerReadPanelFactory
        Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory

        Public Event FormRegionInitializing As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler

        Private _Manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest


        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Public Sub New()
            Me._Manifest = Globals.Factory.CreateFormRegionManifest()
            SysAdminManagerReadPanel.InitializeManifest(Me._Manifest, Globals.Factory)
        End Sub

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Manifest() As Microsoft.Office.Tools.Outlook.FormRegionManifest Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Manifest
            Get
                Return Me._Manifest
            End Get
        End Property

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function CreateFormRegion(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion) As Microsoft.Office.Tools.Outlook.IFormRegion Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion
            Dim form As SysAdminManagerReadPanel = New SysAdminManagerReadPanel(formRegion)
            form.Factory = Me
            Return form
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function GetFormRegionStorage(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Byte() Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage
            Throw New System.NotSupportedException()
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function IsDisplayedForItem(ByVal outlookItem As Object, ByVal formRegionMode As Microsoft.Office.Interop.Outlook.OlFormRegionMode, ByVal formRegionSize As Microsoft.Office.Interop.Outlook.OlFormRegionSize) As Boolean Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem
            Dim cancelArgs As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, False)
            cancelArgs.Cancel = False
            RaiseEvent FormRegionInitializing(Me, cancelArgs)
            Return Not cancelArgs.Cancel
        End Function

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Kind() As Microsoft.Office.Tools.Outlook.FormRegionKindConstants Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            Get
                Return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms
            End Get
        End Property
    End Class
End Class

Partial Class WindowFormRegionCollection

    Friend ReadOnly Property SysAdminManagerReadPanel() As SysAdminManagerReadPanel
        Get
            For Each Item As Object In Me
                If (TypeOf (Item) Is SysAdminManagerReadPanel) Then
                    Return Item
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class