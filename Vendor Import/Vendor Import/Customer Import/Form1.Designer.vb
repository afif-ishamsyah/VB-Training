<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VendorImport
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
        Me.CancelButtons = New System.Windows.Forms.Button()
        Me.DatabaseBox = New System.Windows.Forms.ComboBox()
        Me.UploadButton = New System.Windows.Forms.Button()
        Me.FileNameTextbox = New System.Windows.Forms.TextBox()
        Me.DatabaseLabel = New System.Windows.Forms.Label()
        Me.FileLabel = New System.Windows.Forms.Label()
        Me.CheckExistButton = New System.Windows.Forms.Button()
        Me.FirstCharacterTextBox = New System.Windows.Forms.TextBox()
        Me.FirstCharacterLabel = New System.Windows.Forms.Label()
        Me.ResultLabel = New System.Windows.Forms.Label()
        Me.ResultComboBox = New System.Windows.Forms.ComboBox()
        Me.SearchIDButton = New System.Windows.Forms.Button()
        Me.SearchIDGroup = New System.Windows.Forms.GroupBox()
        Me.SearchNameGroup = New System.Windows.Forms.GroupBox()
        Me.VendorNameListView = New System.Windows.Forms.ListView()
        Me.SearchNameButton = New System.Windows.Forms.Button()
        Me.SearchNameTextBox = New System.Windows.Forms.TextBox()
        Me.SearchNameLabel = New System.Windows.Forms.Label()
        Me.SearchIDGroup.SuspendLayout()
        Me.SearchNameGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'SearchButton
        '
        Me.SearchButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SearchButton.Location = New System.Drawing.Point(362, 44)
        Me.SearchButton.Margin = New System.Windows.Forms.Padding(2)
        Me.SearchButton.Name = "SearchButton"
        Me.SearchButton.Size = New System.Drawing.Size(124, 22)
        Me.SearchButton.TabIndex = 0
        Me.SearchButton.Text = "Search File"
        Me.SearchButton.UseVisualStyleBackColor = True
        '
        'CancelButtons
        '
        Me.CancelButtons.Location = New System.Drawing.Point(478, 273)
        Me.CancelButtons.Margin = New System.Windows.Forms.Padding(2)
        Me.CancelButtons.Name = "CancelButtons"
        Me.CancelButtons.Size = New System.Drawing.Size(118, 34)
        Me.CancelButtons.TabIndex = 1
        Me.CancelButtons.Text = "Cancel"
        Me.CancelButtons.UseVisualStyleBackColor = True
        '
        'DatabaseBox
        '
        Me.DatabaseBox.FormattingEnabled = True
        Me.DatabaseBox.Items.AddRange(New Object() {"LIPDAT - Liputan 6", "KLYBVI - Brilio Ventura Indonesia", "KLYKKI - Kreator Kreatif Indonesia", "KLYKPL - Kapan Lagi", "CMWIDT - Aplikasi Pesan Indonesia", "CMWTRN - Training Only"})
        Me.DatabaseBox.Location = New System.Drawing.Point(66, 12)
        Me.DatabaseBox.Name = "DatabaseBox"
        Me.DatabaseBox.Size = New System.Drawing.Size(291, 21)
        Me.DatabaseBox.TabIndex = 2
        Me.DatabaseBox.Text = "Select Database"
        '
        'UploadButton
        '
        Me.UploadButton.Location = New System.Drawing.Point(362, 273)
        Me.UploadButton.Margin = New System.Windows.Forms.Padding(2)
        Me.UploadButton.Name = "UploadButton"
        Me.UploadButton.Size = New System.Drawing.Size(112, 34)
        Me.UploadButton.TabIndex = 3
        Me.UploadButton.Text = "Upload File"
        Me.UploadButton.UseVisualStyleBackColor = True
        '
        'FileNameTextbox
        '
        Me.FileNameTextbox.Location = New System.Drawing.Point(41, 44)
        Me.FileNameTextbox.Name = "FileNameTextbox"
        Me.FileNameTextbox.Size = New System.Drawing.Size(316, 20)
        Me.FileNameTextbox.TabIndex = 4
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
        'FileLabel
        '
        Me.FileLabel.AutoSize = True
        Me.FileLabel.Location = New System.Drawing.Point(9, 47)
        Me.FileLabel.Name = "FileLabel"
        Me.FileLabel.Size = New System.Drawing.Size(23, 13)
        Me.FileLabel.TabIndex = 7
        Me.FileLabel.Text = "File"
        '
        'CheckExistButton
        '
        Me.CheckExistButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CheckExistButton.Location = New System.Drawing.Point(491, 44)
        Me.CheckExistButton.Name = "CheckExistButton"
        Me.CheckExistButton.Size = New System.Drawing.Size(112, 22)
        Me.CheckExistButton.TabIndex = 8
        Me.CheckExistButton.Text = "Check Exist"
        Me.CheckExistButton.UseVisualStyleBackColor = True
        '
        'FirstCharacterTextBox
        '
        Me.FirstCharacterTextBox.Location = New System.Drawing.Point(116, 26)
        Me.FirstCharacterTextBox.MaxLength = 1
        Me.FirstCharacterTextBox.Name = "FirstCharacterTextBox"
        Me.FirstCharacterTextBox.Size = New System.Drawing.Size(89, 20)
        Me.FirstCharacterTextBox.TabIndex = 0
        '
        'FirstCharacterLabel
        '
        Me.FirstCharacterLabel.AutoSize = True
        Me.FirstCharacterLabel.Location = New System.Drawing.Point(6, 29)
        Me.FirstCharacterLabel.Name = "FirstCharacterLabel"
        Me.FirstCharacterLabel.Size = New System.Drawing.Size(104, 13)
        Me.FirstCharacterLabel.TabIndex = 1
        Me.FirstCharacterLabel.Text = "Insert First Character"
        '
        'ResultLabel
        '
        Me.ResultLabel.AutoSize = True
        Me.ResultLabel.Location = New System.Drawing.Point(6, 89)
        Me.ResultLabel.Name = "ResultLabel"
        Me.ResultLabel.Size = New System.Drawing.Size(37, 13)
        Me.ResultLabel.TabIndex = 3
        Me.ResultLabel.Text = "Result"
        '
        'ResultComboBox
        '
        Me.ResultComboBox.FormattingEnabled = True
        Me.ResultComboBox.Location = New System.Drawing.Point(52, 86)
        Me.ResultComboBox.Name = "ResultComboBox"
        Me.ResultComboBox.Size = New System.Drawing.Size(153, 21)
        Me.ResultComboBox.TabIndex = 4
        '
        'SearchIDButton
        '
        Me.SearchIDButton.Location = New System.Drawing.Point(116, 52)
        Me.SearchIDButton.Name = "SearchIDButton"
        Me.SearchIDButton.Size = New System.Drawing.Size(89, 23)
        Me.SearchIDButton.TabIndex = 5
        Me.SearchIDButton.Text = "Search"
        Me.SearchIDButton.UseVisualStyleBackColor = True
        '
        'SearchIDGroup
        '
        Me.SearchIDGroup.Controls.Add(Me.FirstCharacterLabel)
        Me.SearchIDGroup.Controls.Add(Me.ResultComboBox)
        Me.SearchIDGroup.Controls.Add(Me.SearchIDButton)
        Me.SearchIDGroup.Controls.Add(Me.FirstCharacterTextBox)
        Me.SearchIDGroup.Controls.Add(Me.ResultLabel)
        Me.SearchIDGroup.Location = New System.Drawing.Point(12, 77)
        Me.SearchIDGroup.Name = "SearchIDGroup"
        Me.SearchIDGroup.Size = New System.Drawing.Size(217, 190)
        Me.SearchIDGroup.TabIndex = 9
        Me.SearchIDGroup.TabStop = False
        Me.SearchIDGroup.Text = "Get Vendor ID"
        '
        'SearchNameGroup
        '
        Me.SearchNameGroup.Controls.Add(Me.VendorNameListView)
        Me.SearchNameGroup.Controls.Add(Me.SearchNameButton)
        Me.SearchNameGroup.Controls.Add(Me.SearchNameTextBox)
        Me.SearchNameGroup.Controls.Add(Me.SearchNameLabel)
        Me.SearchNameGroup.Location = New System.Drawing.Point(235, 77)
        Me.SearchNameGroup.Name = "SearchNameGroup"
        Me.SearchNameGroup.Size = New System.Drawing.Size(368, 190)
        Me.SearchNameGroup.TabIndex = 10
        Me.SearchNameGroup.TabStop = False
        Me.SearchNameGroup.Text = "Get Vendor Name"
        '
        'VendorNameListView
        '
        Me.VendorNameListView.Location = New System.Drawing.Point(9, 65)
        Me.VendorNameListView.Name = "VendorNameListView"
        Me.VendorNameListView.Size = New System.Drawing.Size(353, 119)
        Me.VendorNameListView.TabIndex = 3
        Me.VendorNameListView.UseCompatibleStateImageBehavior = False
        '
        'SearchNameButton
        '
        Me.SearchNameButton.Location = New System.Drawing.Point(287, 25)
        Me.SearchNameButton.Name = "SearchNameButton"
        Me.SearchNameButton.Size = New System.Drawing.Size(75, 23)
        Me.SearchNameButton.TabIndex = 2
        Me.SearchNameButton.Text = "Search"
        Me.SearchNameButton.UseVisualStyleBackColor = True
        '
        'SearchNameTextBox
        '
        Me.SearchNameTextBox.Location = New System.Drawing.Point(113, 26)
        Me.SearchNameTextBox.Name = "SearchNameTextBox"
        Me.SearchNameTextBox.Size = New System.Drawing.Size(169, 20)
        Me.SearchNameTextBox.TabIndex = 1
        '
        'SearchNameLabel
        '
        Me.SearchNameLabel.AutoSize = True
        Me.SearchNameLabel.Location = New System.Drawing.Point(6, 29)
        Me.SearchNameLabel.Name = "SearchNameLabel"
        Me.SearchNameLabel.Size = New System.Drawing.Size(101, 13)
        Me.SearchNameLabel.TabIndex = 0
        Me.SearchNameLabel.Text = "Insert Vendor Name"
        '
        'VendorImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(609, 316)
        Me.Controls.Add(Me.SearchNameGroup)
        Me.Controls.Add(Me.SearchIDGroup)
        Me.Controls.Add(Me.CheckExistButton)
        Me.Controls.Add(Me.FileLabel)
        Me.Controls.Add(Me.DatabaseLabel)
        Me.Controls.Add(Me.FileNameTextbox)
        Me.Controls.Add(Me.UploadButton)
        Me.Controls.Add(Me.DatabaseBox)
        Me.Controls.Add(Me.CancelButtons)
        Me.Controls.Add(Me.SearchButton)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "VendorImport"
        Me.Text = "Vendor Import"
        Me.SearchIDGroup.ResumeLayout(False)
        Me.SearchIDGroup.PerformLayout()
        Me.SearchNameGroup.ResumeLayout(False)
        Me.SearchNameGroup.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SearchButton As Button
    Friend WithEvents SearchDialog As OpenFileDialog
    Friend WithEvents CancelButtons As Button
    Friend WithEvents DatabaseBox As ComboBox
    Friend WithEvents UploadButton As Button
    Friend WithEvents FileNameTextbox As TextBox
    Friend WithEvents DatabaseLabel As Label
    Friend WithEvents FileLabel As Label
    Friend WithEvents CheckExistButton As Button
    Friend WithEvents FirstCharacterTextBox As TextBox
    Friend WithEvents FirstCharacterLabel As Label
    Friend WithEvents ResultLabel As Label
    Friend WithEvents ResultComboBox As ComboBox
    Friend WithEvents SearchIDButton As Button
    Friend WithEvents SearchIDGroup As GroupBox
    Friend WithEvents SearchNameGroup As GroupBox
    Friend WithEvents SearchNameTextBox As TextBox
    Friend WithEvents SearchNameLabel As Label
    Friend WithEvents SearchNameButton As Button
    Friend WithEvents VendorNameListView As ListView
End Class
