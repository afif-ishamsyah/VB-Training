<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FileLabel
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
        Me.SearchButton = New System.Windows.Forms.Button()
        Me.SearchDialog = New System.Windows.Forms.OpenFileDialog()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.DatabaseBox = New System.Windows.Forms.ComboBox()
        Me.UploadButton = New System.Windows.Forms.Button()
        Me.FileNameTextbox = New System.Windows.Forms.TextBox()
        Me.LoadingLabel = New System.Windows.Forms.Label()
        Me.DatabaseLabel = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'SearchButton
        '
        Me.SearchButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SearchButton.Location = New System.Drawing.Point(12, 77)
        Me.SearchButton.Margin = New System.Windows.Forms.Padding(2)
        Me.SearchButton.Name = "SearchButton"
        Me.SearchButton.Size = New System.Drawing.Size(112, 41)
        Me.SearchButton.TabIndex = 0
        Me.SearchButton.Text = "Search File"
        Me.SearchButton.UseVisualStyleBackColor = True
        '
        'CancelButton
        '
        Me.CancelButton.Location = New System.Drawing.Point(264, 92)
        Me.CancelButton.Margin = New System.Windows.Forms.Padding(2)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(112, 26)
        Me.CancelButton.TabIndex = 1
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'DatabaseBox
        '
        Me.DatabaseBox.FormattingEnabled = True
        Me.DatabaseBox.Items.AddRange(New Object() {"CMWTRN", "LIPDAT", "KLYBVI", "KLYKKI", "KLYKPL"})
        Me.DatabaseBox.Location = New System.Drawing.Point(66, 12)
        Me.DatabaseBox.Name = "DatabaseBox"
        Me.DatabaseBox.Size = New System.Drawing.Size(185, 21)
        Me.DatabaseBox.TabIndex = 2
        Me.DatabaseBox.Text = "Select Database"
        '
        'UploadButton
        '
        Me.UploadButton.Location = New System.Drawing.Point(139, 77)
        Me.UploadButton.Margin = New System.Windows.Forms.Padding(2)
        Me.UploadButton.Name = "UploadButton"
        Me.UploadButton.Size = New System.Drawing.Size(112, 41)
        Me.UploadButton.TabIndex = 3
        Me.UploadButton.Text = "Upload File"
        Me.UploadButton.UseVisualStyleBackColor = True
        '
        'FileNameTextbox
        '
        Me.FileNameTextbox.Location = New System.Drawing.Point(41, 44)
        Me.FileNameTextbox.Name = "FileNameTextbox"
        Me.FileNameTextbox.Size = New System.Drawing.Size(210, 20)
        Me.FileNameTextbox.TabIndex = 4
        '
        'LoadingLabel
        '
        Me.LoadingLabel.AutoSize = True
        Me.LoadingLabel.BackColor = System.Drawing.SystemColors.Control
        Me.LoadingLabel.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LoadingLabel.Location = New System.Drawing.Point(281, 35)
        Me.LoadingLabel.Name = "LoadingLabel"
        Me.LoadingLabel.Size = New System.Drawing.Size(77, 15)
        Me.LoadingLabel.TabIndex = 5
        Me.LoadingLabel.Text = "PLEASE WAIT"
        '
        'DatabaseLabel
        '
        Me.DatabaseLabel.AutoSize = True
        Me.DatabaseLabel.Location = New System.Drawing.Point(9, 15)
        Me.DatabaseLabel.Name = "DatabaseLabel"
        Me.DatabaseLabel.Size = New System.Drawing.Size(53, 13)
        Me.DatabaseLabel.TabIndex = 6
        Me.DatabaseLabel.Text = "Database"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "File"
        '
        'FileLabel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 129)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DatabaseLabel)
        Me.Controls.Add(Me.LoadingLabel)
        Me.Controls.Add(Me.FileNameTextbox)
        Me.Controls.Add(Me.UploadButton)
        Me.Controls.Add(Me.DatabaseBox)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.SearchButton)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "FileLabel"
        Me.Text = "Customer Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SearchButton As Button
    Friend WithEvents SearchDialog As OpenFileDialog
    Friend WithEvents CancelButton As Button
    Friend WithEvents DatabaseBox As ComboBox
    Friend WithEvents UploadButton As Button
    Friend WithEvents FileNameTextbox As TextBox
    Friend WithEvents LoadingLabel As Label
    Friend WithEvents DatabaseLabel As Label
    Friend WithEvents Label2 As Label
End Class
