Public Class unpack
    Dim unpacker As New Unpacker()

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        unpacker.s_file.Clear()
        unpacker.s_length.Clear()
        Dim a_ = False
        Try
            Label3.Show()
            Label3.ForeColor = Color.Goldenrod
            Label3.Text = "Unpacking '" & IO.Path.GetFileName(TextBox1.Text) & "'..."
            Update()
            unpacker.InitUnpack(TextBox1.Text)
        Catch ex As Exception
            a_ = True
            Label3.ForeColor = Color.Firebrick
            Label3.Text = "Failed to unpack!"
            MessageBox.Show("The unpacked hasn't been able to unpack this file, please make sure it is packed with 'Standalone EXE Document Locker' by 4dots Software" & vbCrLf & vbCrLf & "Exception info:" & vbCrLf & vbCrLf & ex.Message, "Failed to unpack!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        If a_ = False Then
            Label3.ForeColor = Color.ForestGreen
            Label3.Text = "Successfully unpacked!"
            If unpacker.rtfnotused Then
                MessageBox.Show("An alternative method has been used, the unpacked file(s) might need further editing!" & vbCrLf & vbCrLf & "(if the packed exe had a text-based file, it should be fine)", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            If MessageBox.Show("Successfully unpacked!" & vbCrLf & vbCrLf & "Do you want to open the containing folder?", "Unpacked!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                Process.Start("explorer", My.Application.Info.DirectoryPath & "\4dotsEXEDocumentPacker_Extracted\" & IO.Path.GetFileNameWithoutExtension(unpacker.GetDumpedName()) & "\")
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a_ As New OpenFileDialog()
        a_.Filter = "Executable Application|*.exe"
        If a_.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = a_.FileName
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Length = 0 Then
            TextBox1.BackColor = ColorTranslator.FromHtml("#191919")
            Button1.Enabled = False
        Else
            If IO.File.Exists(TextBox1.Text) Then
                TextBox1.BackColor = ColorTranslator.FromHtml("#191919")
                Button1.Enabled = True
            Else
                TextBox1.BackColor = ColorTranslator.FromHtml("#391919")
                Button1.Enabled = False
            End If
        End If
    End Sub

    Private Sub unpack_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
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
End Class