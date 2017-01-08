<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExcelReader
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
        Me.lblSourcePath = New System.Windows.Forms.Label()
        Me.lblDestinationPath = New System.Windows.Forms.Label()
        Me.txtSourcePath = New System.Windows.Forms.TextBox()
        Me.txtDestinationPath = New System.Windows.Forms.TextBox()
        Me.btnSourceBrowse = New System.Windows.Forms.Button()
        Me.btnDestinationBrowse = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblSourcePath
        '
        Me.lblSourcePath.AutoSize = True
        Me.lblSourcePath.Location = New System.Drawing.Point(12, 35)
        Me.lblSourcePath.Name = "lblSourcePath"
        Me.lblSourcePath.Size = New System.Drawing.Size(70, 13)
        Me.lblSourcePath.TabIndex = 0
        Me.lblSourcePath.Text = "Source folder"
        '
        'lblDestinationPath
        '
        Me.lblDestinationPath.AutoSize = True
        Me.lblDestinationPath.Location = New System.Drawing.Point(12, 89)
        Me.lblDestinationPath.Name = "lblDestinationPath"
        Me.lblDestinationPath.Size = New System.Drawing.Size(76, 13)
        Me.lblDestinationPath.TabIndex = 1
        Me.lblDestinationPath.Text = "Destination file"
        '
        'txtSourcePath
        '
        Me.txtSourcePath.Location = New System.Drawing.Point(104, 28)
        Me.txtSourcePath.Name = "txtSourcePath"
        Me.txtSourcePath.Size = New System.Drawing.Size(434, 20)
        Me.txtSourcePath.TabIndex = 2
        '
        'txtDestinationPath
        '
        Me.txtDestinationPath.Location = New System.Drawing.Point(104, 82)
        Me.txtDestinationPath.Name = "txtDestinationPath"
        Me.txtDestinationPath.Size = New System.Drawing.Size(434, 20)
        Me.txtDestinationPath.TabIndex = 3
        '
        'btnSourceBrowse
        '
        Me.btnSourceBrowse.Location = New System.Drawing.Point(560, 24)
        Me.btnSourceBrowse.Name = "btnSourceBrowse"
        Me.btnSourceBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnSourceBrowse.TabIndex = 4
        Me.btnSourceBrowse.Text = "Browse"
        Me.btnSourceBrowse.UseVisualStyleBackColor = True
        '
        'btnDestinationBrowse
        '
        Me.btnDestinationBrowse.Location = New System.Drawing.Point(560, 78)
        Me.btnDestinationBrowse.Name = "btnDestinationBrowse"
        Me.btnDestinationBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnDestinationBrowse.TabIndex = 5
        Me.btnDestinationBrowse.Text = "Browse"
        Me.btnDestinationBrowse.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(235, 125)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 6
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(347, 124)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ExcelReader
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(656, 163)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnDestinationBrowse)
        Me.Controls.Add(Me.btnSourceBrowse)
        Me.Controls.Add(Me.txtDestinationPath)
        Me.Controls.Add(Me.txtSourcePath)
        Me.Controls.Add(Me.lblDestinationPath)
        Me.Controls.Add(Me.lblSourcePath)
        Me.Name = "ExcelReader"
        Me.Text = "Excel Reader"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblSourcePath As Label
    Friend WithEvents lblDestinationPath As Label
    Friend WithEvents txtSourcePath As TextBox
    Friend WithEvents txtDestinationPath As TextBox
    Friend WithEvents btnSourceBrowse As Button
    Friend WithEvents btnDestinationBrowse As Button
    Friend WithEvents btnRun As Button
    Friend WithEvents btnClose As Button
End Class
