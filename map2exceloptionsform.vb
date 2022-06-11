Public Class map2exceloptions

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            setmtckey("options", "notesincomments", "1")
        Else
            setmtckey("options", "notesincomments", "0")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub map2exceloptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub limittovisible_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles limittovisible.CheckedChanged
        If limittovisible.Checked Then
            setmtckey("options", "limittovisible", "1")
        Else
            setmtckey("options", "limittovisible", "0")
        End If
    End Sub

    Private Sub AddTopicHyperlinksCheckbox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddTopicHyperlinksCheckbox.CheckedChanged
        If AddTopicHyperlinksCheckbox.Checked Then
            setmtckey("options", "addtopichyperlinks", "1")
        Else
            setmtckey("options", "addtopichyperlinks", "0")
        End If
    End Sub

    Private Sub AddExternalHyperlinksCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddExternalHyperlinksCheckBox.CheckedChanged
        If AddExternalHyperlinksCheckBox.Checked Then
            setmtckey("options", "addhyperlinks", "1")
        Else
            setmtckey("options", "addhyperlinks", "0")
        End If
    End Sub

    Private Sub AddImagetoCommentCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddImagetoCommentCheckBox.CheckedChanged
        If AddImagetoCommentCheckBox.Checked Then
            setmtckey("options", "addimagetocomment", "1")
        Else
            setmtckey("options", "addimagetocomment", "0")
        End If
    End Sub

    Private Sub AddImagestoCellsCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddImagestoCellsCheckBox.CheckedChanged
        If AddImagestoCellsCheckBox.Checked Then
            setmtckey("options", "addimagetocell", "1")
        Else
            setmtckey("options", "addimagetocell", "0")
        End If
    End Sub


    Private Sub AddOutlineNumbers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddOutlineNumbers.CheckedChanged
        If AddOutlineNumbers.Checked Then
            setmtckey("options", "addoutlinenumbers", "1")
        Else
            setmtckey("options", "addoutlinenumbers", "0")
        End If
    End Sub
   
    Private Sub wraptextcheckbox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles wraptextcheckbox.CheckedChanged
        If wraptextcheckbox.Checked Then
            setmtckey("options", "wraptext", "1")
        Else
            setmtckey("options", "wraptext", "0")
        End If
    End Sub

    Private Sub Label1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles licencekeybox.TextChanged
        setmtckey("options", "key", licencekeybox.Text)
    End Sub


    Private Sub SetTemplateFileNameButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetTemplateFileNameButton.Click
        setmtckey("options", "templatefile", getmapforimport())
        TemplateFileNameLabel.Text = getmtckey("options", "templatefile")
    End Sub

    Private Sub ResetTemplateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetTemplateButton.Click
        setmtckey("options", "templatefile", "")
    End Sub
End Class