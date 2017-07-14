Public Class TextEditorForm

    Private Sub RichTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox.TextChanged
        'Update the status bar
        Dim WordsFound As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(RichTextBox.Text, "[^\p{L}]*\p{L}+[']?\p{L}*[^\p{L}]*")
        StatusLabelWordCount.Text = "Word Count: " & WordsFound.Count

        'Update the about page
        Dim WordCountDictionary As New Dictionary(Of String, Integer)

        For ThisWord As Integer = 0 To WordsFound.Count - 1

            Dim CurrentWordFound As String = WordsFound.Item(ThisWord).Value.ToUpper()
            CurrentWordFound = System.Text.RegularExpressions.Regex.Replace(CurrentWordFound, "[^A-Za-z'-]+", String.Empty)

            If WordCountDictionary.ContainsKey(CurrentWordFound) Then
                WordCountDictionary(CurrentWordFound) += 1
            Else
                WordCountDictionary.Add(CurrentWordFound, 1)
            End If

        Next

        DataGridView.Rows.Clear()

        For Each ThisWord As String In WordCountDictionary.Keys
            DataGridView.Rows.Add(WordCountDictionary(ThisWord), ThisWord)
        Next

    End Sub

    Private Sub Me_Exit(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.FormClosing
        If Not String.IsNullOrEmpty(RichTextBox.Text) Then
            Dim Confirmation As Integer = MessageBox.Show("Would you like to save changes made to " & Me.Text.Replace(" - Text Editor", "?"), "Closing - Text Editor", MessageBoxButtons.YesNoCancel)
            If Confirmation = DialogResult.Yes Then
                SaveToolStripMenuItem_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim Confirmation As Integer = MessageBox.Show("Would you like to save changes made to " & Me.Text.Replace(" - Text Editor", "?"), "Creating a new file - Text Editor", MessageBoxButtons.YesNoCancel)
        If Confirmation = DialogResult.Yes Then
            SaveToolStripMenuItem_Click(sender, e)
            Confirmation = DialogResult.No
        End If
        If Confirmation = DialogResult.No Then
            RichTextBox.Text = ""
            Me.Text = "Untitled - Text Editor"
        End If
    End Sub

    Private Sub LoadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click
        Dim Confirmation As Integer = MessageBox.Show("Would you like to save changes made to " & Me.Text.Replace(" - Text Editor", "?"), "Opening a new file - Text Editor", MessageBoxButtons.YesNoCancel)
        If Confirmation = DialogResult.Yes Then
            SaveToolStripMenuItem_Click(sender, e)
            Confirmation = DialogResult.No
        End If
        If Confirmation = DialogResult.No Then
            OpenFileDialog.Filter = "Text Documents (*.txt)|*.txt"
            OpenFileDialog.FileName = "*.txt"
            If OpenFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                RichTextBox.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog.FileName)
                Me.Text = IO.Path.GetFileName(OpenFileDialog.FileName).Replace(".txt", "") & " - Text Editor"
            End If
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveFileDialog.Filter = "Text Documents (*.txt)|*.txt"
        SaveFileDialog.FileName = "*.txt"
        If SaveFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            My.Computer.FileSystem.WriteAllText(SaveFileDialog.FileName, RichTextBox.Text, False)
            Me.Text = IO.Path.GetFileName(SaveFileDialog.FileName).Replace(".txt", "") & " - Text Editor"
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox.Paste()
    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        If FontDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox.Font = FontDialog.Font
            DataGridView.Font = FontDialog.Font
        End If
    End Sub

    Private Sub BackgroundColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackgroundColorToolStripMenuItem.Click
        If BackgroundColorDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox.BackColor = BackgroundColorDialog.Color
            DataGridView.RowsDefaultCellStyle.BackColor = BackgroundColorDialog.Color
            DataGridView.BackgroundColor = BackgroundColorDialog.Color
        End If
    End Sub

    Private Sub TextColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextColorToolStripMenuItem.Click
        If TextColorDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            RichTextBox.ForeColor = TextColorDialog.Color
            DataGridView.ForeColor = TextColorDialog.Color
        End If
    End Sub

    Private Sub ToggleNightshadeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToggleNightshadeToolStripMenuItem.Click
        If ToggleNightshadeToolStripMenuItem.Text = "Turn Night Styling Off" Then
            RichTextBox.ForeColor = Color.Black
            RichTextBox.BackColor = Color.White
            DataGridView.ForeColor = Color.Black
            DataGridView.RowsDefaultCellStyle.BackColor = Color.White
            DataGridView.BackgroundColor = Color.White
            EraseHightlightsToolStripMenuItem_Click(sender, e)
            ToggleNightshadeToolStripMenuItem.Text = "Turn Night Styling On"
        Else
            RichTextBox.ForeColor = Color.FromArgb(120, 120, 120)
            RichTextBox.BackColor = Color.FromArgb(51, 51, 51)
            DataGridView.ForeColor = Color.FromArgb(120, 120, 120)
            DataGridView.RowsDefaultCellStyle.BackColor = Color.FromArgb(51, 51, 51)
            DataGridView.BackgroundColor = Color.FromArgb(51, 51, 51)
            EraseHightlightsToolStripMenuItem_Click(sender, e)
            ToggleNightshadeToolStripMenuItem.Text = "Turn Night Styling Off"
        End If
    End Sub

    Private Sub HighlightWordsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HighlightWordsToolStripMenuItem.Click
        Dim SearchInput As String = InputBox("", "Search - Text Editor")
        Dim SearchPosition As Integer = 0
        Dim NumberOfMatches As Integer = 0
        If SearchInput.Length > 0 Then
            While RichTextBox.Find(SearchInput, SearchPosition, RichTextBoxFinds.WholeWord) <> -1
                RichTextBox.SelectionBackColor = Color.Yellow
                SearchPosition = RichTextBox.Find(SearchInput, SearchPosition, RichTextBoxFinds.WholeWord) + SearchInput.Length
                NumberOfMatches = NumberOfMatches + 1
                If RichTextBox.Find(SearchInput, SearchPosition, RichTextBoxFinds.WholeWord) = 0 Then
                    Exit While
                End If
            End While
            RichTextBox.DeselectAll()
            If NumberOfMatches = 1 Then
                MsgBox("Found " & NumberOfMatches & " match.")
            Else
                MsgBox("Found " & NumberOfMatches & " matches.")
            End If
        End If
    End Sub

    Private Sub EraseHightlightsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EraseHightlightsToolStripMenuItem.Click
        RichTextBox.SelectAll()
        RichTextBox.SelectionBackColor = RichTextBox.BackColor
        RichTextBox.DeselectAll()
    End Sub
End Class
