<System.ComponentModel.ToolboxItemAttribute(False)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SysAdminManagerWritePanel
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
        manifest.FormRegionName = "SysAdminManagerWritePanel"
        manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining
        manifest.ShowInspectorRead = False
        manifest.ShowReadingPane = False

    End Sub

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DelayLabel = New System.Windows.Forms.Label()
        Me.EffortTypeLabel = New System.Windows.Forms.Label()
        Me.OriginLabel = New System.Windows.Forms.Label()
        Me.EffortTypeComboBox = New System.Windows.Forms.ComboBox()
        Me.OriginComboBox = New System.Windows.Forms.ComboBox()
        Me.DelayerDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.SendImmediatlyButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DelayLabel
        '
        Me.DelayLabel.AutoSize = True
        Me.DelayLabel.Location = New System.Drawing.Point(3, 22)
        Me.DelayLabel.Name = "DelayLabel"
        Me.DelayLabel.Size = New System.Drawing.Size(34, 13)
        Me.DelayLabel.TabIndex = 11
        Me.DelayLabel.Text = "Delay"
        '
        'EffortTypeLabel
        '
        Me.EffortTypeLabel.AutoSize = True
        Me.EffortTypeLabel.Location = New System.Drawing.Point(701, 22)
        Me.EffortTypeLabel.Name = "EffortTypeLabel"
        Me.EffortTypeLabel.Size = New System.Drawing.Size(59, 13)
        Me.EffortTypeLabel.TabIndex = 10
        Me.EffortTypeLabel.Text = "Effort Type"
        '
        'OriginLabel
        '
        Me.OriginLabel.AutoSize = True
        Me.OriginLabel.Location = New System.Drawing.Point(434, 22)
        Me.OriginLabel.Name = "OriginLabel"
        Me.OriginLabel.Size = New System.Drawing.Size(34, 13)
        Me.OriginLabel.TabIndex = 9
        Me.OriginLabel.Text = "Origin"
        '
        'EffortTypeComboBox
        '
        Me.EffortTypeComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.EffortTypeComboBox.FormattingEnabled = True
        Me.EffortTypeComboBox.Items.AddRange(New Object() {"Quick Request", "Long Request", "Emergency"})
        Me.EffortTypeComboBox.Location = New System.Drawing.Point(765, 18)
        Me.EffortTypeComboBox.Name = "EffortTypeComboBox"
        Me.EffortTypeComboBox.Size = New System.Drawing.Size(200, 21)
        Me.EffortTypeComboBox.TabIndex = 8
        '
        'OriginComboBox
        '
        Me.OriginComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.OriginComboBox.FormattingEnabled = True
        Me.OriginComboBox.Items.AddRange(New Object() {"Project Task", "Emergency Request", "Incident", "Incident - alert", "Incident - outage"})
        Me.OriginComboBox.Location = New System.Drawing.Point(474, 18)
        Me.OriginComboBox.Name = "OriginComboBox"
        Me.OriginComboBox.Size = New System.Drawing.Size(200, 21)
        Me.OriginComboBox.TabIndex = 7
        '
        'DelayerDateTimePicker
        '
        Me.DelayerDateTimePicker.CustomFormat = "ddd yyyy-MM-dd HH:mm"
        Me.DelayerDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DelayerDateTimePicker.Location = New System.Drawing.Point(43, 18)
        Me.DelayerDateTimePicker.Name = "DelayerDateTimePicker"
        Me.DelayerDateTimePicker.ShowCheckBox = True
        Me.DelayerDateTimePicker.Size = New System.Drawing.Size(200, 20)
        Me.DelayerDateTimePicker.TabIndex = 6
        '
        'SendImmediatlyButton
        '
        Me.SendImmediatlyButton.BackColor = System.Drawing.Color.Yellow
        Me.SendImmediatlyButton.Location = New System.Drawing.Point(266, 17)
        Me.SendImmediatlyButton.Name = "SendImmediatlyButton"
        Me.SendImmediatlyButton.Size = New System.Drawing.Size(107, 23)
        Me.SendImmediatlyButton.TabIndex = 12
        Me.SendImmediatlyButton.Text = "Send Immediatly"
        Me.SendImmediatlyButton.UseVisualStyleBackColor = False
        '
        'SysAdminManagerWritePanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SendImmediatlyButton)
        Me.Controls.Add(Me.DelayLabel)
        Me.Controls.Add(Me.EffortTypeLabel)
        Me.Controls.Add(Me.OriginLabel)
        Me.Controls.Add(Me.EffortTypeComboBox)
        Me.Controls.Add(Me.OriginComboBox)
        Me.Controls.Add(Me.DelayerDateTimePicker)
        Me.Name = "SysAdminManagerWritePanel"
        Me.Size = New System.Drawing.Size(982, 54)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DelayLabel As Windows.Forms.Label
    Friend WithEvents EffortTypeLabel As Windows.Forms.Label
    Friend WithEvents OriginLabel As Windows.Forms.Label
    Friend WithEvents EffortTypeComboBox As Windows.Forms.ComboBox
    Friend WithEvents OriginComboBox As Windows.Forms.ComboBox
    Friend WithEvents DelayerDateTimePicker As Windows.Forms.DateTimePicker
    Friend WithEvents SendImmediatlyButton As Windows.Forms.Button

    Partial Public Class SysAdminManagerWritePanelFactory
        Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory

        Public Event FormRegionInitializing As Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler

        Private _Manifest As Microsoft.Office.Tools.Outlook.FormRegionManifest


        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Public Sub New()
            Me._Manifest = Globals.Factory.CreateFormRegionManifest()
            SysAdminManagerWritePanel.InitializeManifest(Me._Manifest, Globals.Factory)
        End Sub

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        ReadOnly Property Manifest() As Microsoft.Office.Tools.Outlook.FormRegionManifest Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.Manifest
            Get
                Return Me._Manifest
            End Get
        End Property

        <System.Diagnostics.DebuggerNonUserCodeAttribute()>
        Function CreateFormRegion(ByVal formRegion As Microsoft.Office.Interop.Outlook.FormRegion) As Microsoft.Office.Tools.Outlook.IFormRegion Implements Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion
            Dim form As SysAdminManagerWritePanel = New SysAdminManagerWritePanel(formRegion)
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

    Friend ReadOnly Property SysAdminManagerWritePanel() As SysAdminManagerWritePanel
        Get
            For Each Item As Object In Me
                If (TypeOf (Item) Is SysAdminManagerWritePanel) Then
                    Return Item
                End If
            Next
            Return Nothing
        End Get
    End Property
End Class