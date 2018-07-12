<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.TimerScrollGrid = New System.Windows.Forms.Timer(Me.components)
        Me.DataGridView = New System.Windows.Forms.DataGridView()
        Me.FileSystemWatcher = New System.IO.FileSystemWatcher()
        Me.LabelStatus = New System.Windows.Forms.Label()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon
        '
        Me.NotifyIcon.Text = "NotifyIcon1"
        Me.NotifyIcon.Visible = True
        '
        'TimerScrollGrid
        '
        '
        'DataGridView
        '
        Me.DataGridView.AllowUserToAddRows = False
        Me.DataGridView.AllowUserToDeleteRows = False
        Me.DataGridView.AllowUserToResizeColumns = False
        Me.DataGridView.AllowUserToResizeRows = False
        Me.DataGridView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView.CausesValidation = False
        Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView.Location = New System.Drawing.Point(27, 84)
        Me.DataGridView.MultiSelect = False
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.RowTemplate.Height = 24
        Me.DataGridView.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView.Size = New System.Drawing.Size(240, 150)
        Me.DataGridView.TabIndex = 4
        Me.DataGridView.TabStop = False
        '
        'FileSystemWatcher
        '
        Me.FileSystemWatcher.EnableRaisingEvents = True
        Me.FileSystemWatcher.SynchronizingObject = Me
        '
        'LabelStatus
        '
        Me.LabelStatus.AutoSize = True
        Me.LabelStatus.Location = New System.Drawing.Point(24, 9)
        Me.LabelStatus.Name = "LabelStatus"
        Me.LabelStatus.Size = New System.Drawing.Size(83, 17)
        Me.LabelStatus.TabIndex = 1
        Me.LabelStatus.Text = "LabelStatus"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 268)
        Me.Controls.Add(Me.LabelStatus)
        Me.Controls.Add(Me.DataGridView)
        Me.Name = "frmMain"
        Me.Text = "InfoMerc"
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents NotifyIcon As NotifyIcon
    Friend WithEvents TimerScrollGrid As Timer
    Friend WithEvents DataGridView As DataGridView
    Friend WithEvents FileSystemWatcher As IO.FileSystemWatcher
    Friend WithEvents LabelStatus As Label
End Class
