<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class map2exceloptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(map2exceloptions))
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.limittovisible = New System.Windows.Forms.CheckBox()
        Me.AddTopicHyperlinksCheckbox = New System.Windows.Forms.CheckBox()
        Me.AddExternalHyperlinksCheckBox = New System.Windows.Forms.CheckBox()
        Me.AddImagetoCommentCheckBox = New System.Windows.Forms.CheckBox()
        Me.AddImagestoCellsCheckBox = New System.Windows.Forms.CheckBox()
        Me.AddOutlineNumbers = New System.Windows.Forms.CheckBox()
        Me.wraptextcheckbox = New System.Windows.Forms.CheckBox()
        Me.licencekeybox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SetTemplateFileNameButton = New System.Windows.Forms.Button()
        Me.TemplateFileNameLabel = New System.Windows.Forms.Label()
        Me.ResetTemplateButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(12, 28)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(136, 17)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Put Notes in Comments"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 316)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'limittovisible
        '
        Me.limittovisible.AutoSize = True
        Me.limittovisible.Location = New System.Drawing.Point(13, 51)
        Me.limittovisible.Name = "limittovisible"
        Me.limittovisible.Size = New System.Drawing.Size(122, 17)
        Me.limittovisible.TabIndex = 3
        Me.limittovisible.Text = "Limit to visible topics"
        Me.limittovisible.UseVisualStyleBackColor = True
        '
        'AddTopicHyperlinksCheckbox
        '
        Me.AddTopicHyperlinksCheckbox.AutoSize = True
        Me.AddTopicHyperlinksCheckbox.Location = New System.Drawing.Point(13, 74)
        Me.AddTopicHyperlinksCheckbox.Name = "AddTopicHyperlinksCheckbox"
        Me.AddTopicHyperlinksCheckbox.Size = New System.Drawing.Size(127, 17)
        Me.AddTopicHyperlinksCheckbox.TabIndex = 4
        Me.AddTopicHyperlinksCheckbox.Text = "Add Topic Hyperlinks"
        Me.AddTopicHyperlinksCheckbox.UseVisualStyleBackColor = True
        '
        'AddExternalHyperlinksCheckBox
        '
        Me.AddExternalHyperlinksCheckBox.AutoSize = True
        Me.AddExternalHyperlinksCheckBox.Location = New System.Drawing.Point(13, 97)
        Me.AddExternalHyperlinksCheckBox.Name = "AddExternalHyperlinksCheckBox"
        Me.AddExternalHyperlinksCheckBox.Size = New System.Drawing.Size(135, 17)
        Me.AddExternalHyperlinksCheckBox.TabIndex = 5
        Me.AddExternalHyperlinksCheckBox.Text = "Add external hyperlinks"
        Me.AddExternalHyperlinksCheckBox.UseVisualStyleBackColor = True
        '
        'AddImagetoCommentCheckBox
        '
        Me.AddImagetoCommentCheckBox.AutoSize = True
        Me.AddImagetoCommentCheckBox.Location = New System.Drawing.Point(13, 121)
        Me.AddImagetoCommentCheckBox.Name = "AddImagetoCommentCheckBox"
        Me.AddImagetoCommentCheckBox.Size = New System.Drawing.Size(145, 17)
        Me.AddImagetoCommentCheckBox.TabIndex = 6
        Me.AddImagetoCommentCheckBox.Text = "Add Images to comments"
        Me.AddImagetoCommentCheckBox.UseVisualStyleBackColor = True
        '
        'AddImagestoCellsCheckBox
        '
        Me.AddImagestoCellsCheckBox.AutoSize = True
        Me.AddImagestoCellsCheckBox.Location = New System.Drawing.Point(13, 144)
        Me.AddImagestoCellsCheckBox.Name = "AddImagestoCellsCheckBox"
        Me.AddImagestoCellsCheckBox.Size = New System.Drawing.Size(117, 17)
        Me.AddImagestoCellsCheckBox.TabIndex = 7
        Me.AddImagestoCellsCheckBox.Text = "Add images to cells"
        Me.AddImagestoCellsCheckBox.UseVisualStyleBackColor = True
        '
        'AddOutlineNumbers
        '
        Me.AddOutlineNumbers.AutoSize = True
        Me.AddOutlineNumbers.Location = New System.Drawing.Point(13, 168)
        Me.AddOutlineNumbers.Name = "AddOutlineNumbers"
        Me.AddOutlineNumbers.Size = New System.Drawing.Size(126, 17)
        Me.AddOutlineNumbers.TabIndex = 8
        Me.AddOutlineNumbers.Text = "Add Outline Numbers"
        Me.AddOutlineNumbers.UseVisualStyleBackColor = True
        '
        'wraptextcheckbox
        '
        Me.wraptextcheckbox.AutoSize = True
        Me.wraptextcheckbox.Location = New System.Drawing.Point(13, 192)
        Me.wraptextcheckbox.Name = "wraptextcheckbox"
        Me.wraptextcheckbox.Size = New System.Drawing.Size(104, 17)
        Me.wraptextcheckbox.TabIndex = 9
        Me.wraptextcheckbox.Text = "wrap text in cells"
        Me.wraptextcheckbox.UseVisualStyleBackColor = True
        '
        'licencekeybox
        '
        Me.licencekeybox.Location = New System.Drawing.Point(104, 283)
        Me.licencekeybox.Name = "licencekeybox"
        Me.licencekeybox.Size = New System.Drawing.Size(100, 20)
        Me.licencekeybox.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 286)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Licence Key"
        '
        'SetTemplateFileNameButton
        '
        Me.SetTemplateFileNameButton.Location = New System.Drawing.Point(15, 216)
        Me.SetTemplateFileNameButton.Name = "SetTemplateFileNameButton"
        Me.SetTemplateFileNameButton.Size = New System.Drawing.Size(120, 23)
        Me.SetTemplateFileNameButton.TabIndex = 12
        Me.SetTemplateFileNameButton.Text = "Set Template File"
        Me.SetTemplateFileNameButton.UseVisualStyleBackColor = True
        '
        'TemplateFileNameLabel
        '
        Me.TemplateFileNameLabel.AutoSize = True
        Me.TemplateFileNameLabel.Location = New System.Drawing.Point(156, 219)
        Me.TemplateFileNameLabel.Name = "TemplateFileNameLabel"
        Me.TemplateFileNameLabel.Size = New System.Drawing.Size(111, 13)
        Me.TemplateFileNameLabel.TabIndex = 13
        Me.TemplateFileNameLabel.Text = "Template Not Defined"
        '
        'ResetTemplateButton
        '
        Me.ResetTemplateButton.Location = New System.Drawing.Point(15, 245)
        Me.ResetTemplateButton.Name = "ResetTemplateButton"
        Me.ResetTemplateButton.Size = New System.Drawing.Size(115, 23)
        Me.ResetTemplateButton.TabIndex = 14
        Me.ResetTemplateButton.Text = "Reset Template File"
        Me.ResetTemplateButton.UseVisualStyleBackColor = True
        '
        'map2exceloptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(668, 392)
        Me.Controls.Add(Me.ResetTemplateButton)
        Me.Controls.Add(Me.TemplateFileNameLabel)
        Me.Controls.Add(Me.SetTemplateFileNameButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.licencekeybox)
        Me.Controls.Add(Me.wraptextcheckbox)
        Me.Controls.Add(Me.AddOutlineNumbers)
        Me.Controls.Add(Me.AddImagestoCellsCheckBox)
        Me.Controls.Add(Me.AddImagetoCommentCheckBox)
        Me.Controls.Add(Me.AddExternalHyperlinksCheckBox)
        Me.Controls.Add(Me.AddTopicHyperlinksCheckbox)
        Me.Controls.Add(Me.limittovisible)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "map2exceloptions"
        Me.Text = "Map2Excel Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents limittovisible As System.Windows.Forms.CheckBox
    Friend WithEvents AddTopicHyperlinksCheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents AddExternalHyperlinksCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents AddImagetoCommentCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents AddImagestoCellsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents AddOutlineNumbers As System.Windows.Forms.CheckBox
    Friend WithEvents wraptextcheckbox As System.Windows.Forms.CheckBox
    Friend WithEvents licencekeybox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SetTemplateFileNameButton As System.Windows.Forms.Button
    Friend WithEvents TemplateFileNameLabel As System.Windows.Forms.Label
    Friend WithEvents ResetTemplateButton As System.Windows.Forms.Button
End Class
