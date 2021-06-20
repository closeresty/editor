Imports System.IO
Imports System.Drawing.Printing

Public Class frmEditor
    Public writefile As StreamWriter
    Public readfile As StreamReader
    Dim result As DialogResult
    Dim s As Char
    Dim l As Integer
    Dim fpath, fname As String
    Dim slposition As Integer
    Dim msgresult As MsgBoxResult
    Dim saveextfile As String
    Private Function PreparePrintDocument() As PrintDocument
        Dim Print_Document As New PrintDocument
        AddHandler Print_Document.PrintPage, AddressOf Print_PrintPage
        Return Print_Document
    End Function

    Private Sub Print_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim Page As String
        Page = Me.rtbEditor.Text
        e.Graphics.DrawString(Page, rtbEditor.Font, Brushes.Black, 50, 50)
        e.HasMorePages = False
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Try
            If l = 0 Then
                OpenFileDialog1.InitialDirectory = "C:\"
                OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                result = OpenFileDialog1.ShowDialog
                If result = DialogResult.OK Then
                    readfile = New StreamReader(OpenFileDialog1.FileName)
                    rtbEditor.Text = readfile.ReadToEnd()
                    fpath = OpenFileDialog1.FileName()
                    slposition = fpath.LastIndexOf("\")
                    fname = fpath.Substring(slposition + 1)
                    Me.Text = fname
                End If
                
            ElseIf l = 1 Then
        msgresult = MessageBox.Show("Do you want to save?", "Save File", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
        If msgresult = MsgBoxResult.Yes Then
            SaveFileDialog1.InitialDirectory = "C:\"
            SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            result = SaveFileDialog1.ShowDialog
            If result = DialogResult.OK Then
                writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                writefile.Write(rtbEditor.Text)
                writefile.Close()
                        rtbEditor.Text = ""
                        Me.Text = "Untitled"
                        l = 0
            End If
        ElseIf msgresult = MsgBoxResult.No Then
            rtbEditor.Text = ""
            rtbEditor.Focus()
            OpenFileDialog1.InitialDirectory = "C:\"
            OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            result = OpenFileDialog1.ShowDialog
            If result = DialogResult.OK Then
                readfile = New StreamReader(OpenFileDialog1.FileName)
                rtbEditor.Text = readfile.ReadToEnd()
                fpath = OpenFileDialog1.FileName()
                slposition = fpath.LastIndexOf("\")
                fname = fpath.Substring(slposition + 1)
                Me.Text = fname
                        readfile.Close()
                        l = 0
            End If
        End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Try
            SaveFileDialog1.InitialDirectory = "C:\"
            SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            result = SaveFileDialog1.ShowDialog
            If result = DialogResult.OK Then
                writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                writefile.Write(rtbEditor.Text)
                fpath = SaveFileDialog1.FileName()
                slposition = fpath.LastIndexOf("\")
                fname = fpath.Substring(slposition + 1)
                Me.Text = fname
                writefile.Close()
                rtbEditor.Text = ""
                Me.Text = "Untitled"
                l = 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub frmEditor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Untitled"
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If rtbEditor.SelectionLength = 0 Then
            CutToolStripMenuItem.Enabled = False
            CopyToolStripMenuItem.Enabled = False
        Else
            CutToolStripMenuItem.Enabled = True
            CopyToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText(rtbEditor.SelectedText)
        rtbEditor.SelectedText = ""
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        Clipboard.SetText(rtbEditor.SelectedText)
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        rtbEditor.SelectedText = Clipboard.GetText
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetText(rtbEditor.SelectedText)
        rtbEditor.SelectedText = ""
    End Sub

    Private Sub BackcolorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackcolorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            rtbEditor.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub FontcolorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontcolorToolStripMenuItem.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            rtbEditor.SelectionColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            rtbEditor.Font = FontDialog1.Font
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        rtbEditor.SelectionStart = 0
        rtbEditor.SelectionLength = Len(rtbEditor.Text)
    End Sub

    Private Sub rtbEditor_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles rtbEditor.KeyPress
        l = 1
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Try
            If l = 0 Then
                rtbEditor.Text = ""
                rtbEditor.Focus()
                Me.Text = "Untitled"
                l = 0
            ElseIf l = 1 Then
                msgresult = MessageBox.Show("Do you want to save?", "Save File", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
                If msgresult = MsgBoxResult.Yes Then
                    SaveFileDialog1.InitialDirectory = "C:\"
                    SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    result = SaveFileDialog1.ShowDialog
                    If result = DialogResult.OK Then
                        writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                        writefile.Write(rtbEditor.Text)
                        writefile.Close()
                        rtbEditor.Text = ""
                        rtbEditor.Focus()
                        Me.Text = "Untitled"
                        l = 0
                    End If
                ElseIf msgresult = MsgBoxResult.No Then
                    rtbEditor.Text = ""
                    rtbEditor.Focus()
                    Me.Text = "Untitled"
                    l = 0
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
         Dim Print_Preview As New PrintPreviewDialog
        Print_Preview.Document = PreparePrintDocument()
        Print_Preview.WindowState = FormWindowState.Maximized
        Print_Preview.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Try
            If l = 0 Then
                Application.Exit()
            ElseIf l = 1 Then
                msgresult = MessageBox.Show("Do you want to save?", "Save File", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
                If msgresult = MsgBoxResult.Yes Then
                    SaveFileDialog1.InitialDirectory = "C:\"
                    SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    result = SaveFileDialog1.ShowDialog
                    If result = DialogResult.OK Then
                        writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                        writefile.Write(rtbEditor.Text)
                        writefile.Close()
                        Application.Exit()
                    End If
                ElseIf msgresult = MsgBoxResult.No Then
                    Application.Exit()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim Print_Dialog As New PrintDialog
        Print_Dialog.Document = PreparePrintDocument()
        Print_Dialog.ShowDialog()
        If Print_Dialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Print_Dialog.Document.Print()
        End If
    End Sub

    Private Sub NewToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripButton.Click
        Try
            If l = 0 Then
                rtbEditor.Text = ""
                rtbEditor.Focus()
                Me.Text = "Untitled"
                l = 0
            ElseIf l = 1 Then
                msgresult = MessageBox.Show("Do you want to save?", "Save File", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
                If msgresult = MsgBoxResult.Yes Then
                    SaveFileDialog1.InitialDirectory = "C:\"
                    SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    result = SaveFileDialog1.ShowDialog
                    If result = DialogResult.OK Then
                        writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                        writefile.Write(rtbEditor.Text)
                        writefile.Close()
                        rtbEditor.Text = ""
                        rtbEditor.Focus()
                        Me.Text = "Untitled"
                        l = 0
                    End If
                ElseIf msgresult = MsgBoxResult.No Then
                    rtbEditor.Text = ""
                    rtbEditor.Focus()
                    Me.Text = "Untitled"
                    l = 0
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click
        Try
            If l = 0 Then
                OpenFileDialog1.InitialDirectory = "C:\"
                OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                result = OpenFileDialog1.ShowDialog
                If result = DialogResult.OK Then
                    readfile = New StreamReader(OpenFileDialog1.FileName)
                    rtbEditor.Text = readfile.ReadToEnd()
                    fpath = OpenFileDialog1.FileName()
                    slposition = fpath.LastIndexOf("\")
                    fname = fpath.Substring(slposition + 1)
                    Me.Text = fname
                End If

            ElseIf l = 1 Then
                msgresult = MessageBox.Show("Do you want to save?", "Save File", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)
                If msgresult = MsgBoxResult.Yes Then
                    SaveFileDialog1.InitialDirectory = "C:\"
                    SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    result = SaveFileDialog1.ShowDialog
                    If result = DialogResult.OK Then
                        writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                        writefile.Write(rtbEditor.Text)
                        writefile.Close()
                        rtbEditor.Text = ""
                        Me.Text = "Untitled"
                        l = 0
                    End If
                ElseIf msgresult = MsgBoxResult.No Then
                    rtbEditor.Text = ""
                    rtbEditor.Focus()
                    OpenFileDialog1.InitialDirectory = "C:\"
                    OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    result = OpenFileDialog1.ShowDialog
                    If result = DialogResult.OK Then
                        readfile = New StreamReader(OpenFileDialog1.FileName)
                        rtbEditor.Text = readfile.ReadToEnd()
                        fpath = OpenFileDialog1.FileName()
                        slposition = fpath.LastIndexOf("\")
                        fname = fpath.Substring(slposition + 1)
                        Me.Text = fname
                        readfile.Close()
                        l = 0
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        Try
            SaveFileDialog1.InitialDirectory = "C:\"
            SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            result = SaveFileDialog1.ShowDialog
            If result = DialogResult.OK Then
                writefile = New StreamWriter(SaveFileDialog1.FileName, False)
                writefile.Write(rtbEditor.Text)
                fpath = SaveFileDialog1.FileName()
                slposition = fpath.LastIndexOf("\")
                fname = fpath.Substring(slposition + 1)
                Me.Text = fname
                writefile.Close()
                rtbEditor.Text = ""
                Me.Text = "Untitled"
                l = 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripButton.Click
        Dim Print_Dialog As New PrintDialog
        Print_Dialog.Document = PreparePrintDocument()
        Print_Dialog.ShowDialog()
        If Print_Dialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            Print_Dialog.Document.Print()
        End If
    End Sub

    Private Sub CutToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripButton.Click
        Clipboard.Clear()
        Clipboard.SetText(rtbEditor.SelectedText)
        rtbEditor.SelectedText = ""
    End Sub

    Private Sub CopyToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripButton.Click
        Clipboard.SetText(rtbEditor.SelectedText)
    End Sub

    Private Sub PasteToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripButton.Click
        rtbEditor.SelectedText = Clipboard.GetText
    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripButton.Click
        Help.ShowHelp(Me, "F:\vb\Sample Programs\More Programs\MyEditor\Editor_Help.html")
    End Sub

    Private Sub ViewHelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewHelpToolStripMenuItem.Click
        Help.ShowHelp(Me, "F:\vb\Sample Programs\More Programs\MyEditor\Editor_Help.html")
    End Sub
End Class
