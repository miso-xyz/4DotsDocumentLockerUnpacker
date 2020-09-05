Public Class Form1
    Dim name_, password_
    Dim Unpacker As New Unpacker()
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim a_ As New SaveFileDialog()
        a_.Filter = "Text File|*.txt"
        If a_.ShowDialog() = Windows.Forms.DialogResult.OK Then
            IO.File.AppendAllText(a_.FileName, "## 4 Dots (EXE Locker) Unpacker v1.0 - Unpacker by misonothx ##" & vbCrLf & TextBox2.Text)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Clipboard.SetText(TextBox2.Text)
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        str.Show()
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        TextBox1.Text = files(0)
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Length = 0 Then
            TextBox1.BackColor = ColorTranslator.FromHtml("#191919")
        Else
            If IO.File.Exists(TextBox1.Text) Then
                TextBox1.BackColor = ColorTranslator.FromHtml("#191919")
                TextBox2.ForeColor = Color.Goldenrod
                TextBox2.Text = "Decrypting..."
                TextBox2.Update()
                Dim a_ = False
                Try
                    Unpacker.GetNameAndPassword(TextBox1.Text)
                    Unpacker.UnpackFiles(TextBox1.Text, True)
                Catch ex As Exception
                    a_ = True
                    TextBox2.ForeColor = Color.Firebrick
                    TextBox2.Text = "Failed to decrypt!"
                End Try
                If a_ = False Then
                    TextBox2.ForeColor = ColorTranslator.FromHtml("#A4A4A4")
                    TextBox2.Clear()
                    If CheckBox1.Checked Then
                        TextBox2.Text += "Name: " & Unpacker.GetDumpedName() & vbCrLf
                    Else
                        TextBox2.Text += "Name: -" & vbCrLf
                    End If
                    If CheckBox2.Checked Then
                        TextBox2.Text += "Password: " & Unpacker.GetDumpedPassword() & vbCrLf
                    Else
                        TextBox2.Text += "Password: -" & vbCrLf
                    End If
                    TextBox2.Text += vbCrLf & "Files found (" & Unpacker.s_file.Count() & "):"
                    For x = 0 To Unpacker.s_file.Count() - 1
                        If CheckBox4.Checked Then
                            TextBox2.Text += vbCrLf & "- " & Unpacker.s_file(x)
                        Else
                            TextBox2.Text += vbCrLf & "- -"
                        End If
                        TextBox2.Text += " | "
                        If CheckBox3.Checked Then
                            TextBox2.Text += Unpacker.s_length(x)
                        Else
                            TextBox2.Text += "-"
                        End If
                    Next
                End If
                Unpacker.s_file.Clear()
                Unpacker.s_length.Clear()
            Else
                TextBox1.BackColor = ColorTranslator.FromHtml("#391919")
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a_ As New OpenFileDialog()
        a_.Filter = "Executable Application|*.exe"
        If a_.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = a_.FileName
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Length = 0 Then
            Button3.Enabled = False
            LinkLabel1.Enabled = False
        Else
            Button3.Enabled = True
            LinkLabel1.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub
End Class
