Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'error check
        If Not IsNumeric(TextBox1.Text) Then
            TextBox1.Clear()
            MsgBox("Please only enter numbers")
            Exit Sub
        End If
        If Val(TextBox1.Text) <= 0 Then
            TextBox1.Clear()
            MsgBox("Please enter positave values")
            Exit Sub
        End If
        If ListBox1.Items.Count = 10 Then
            TextBox1.Clear()
            MsgBox("No more bills can be entered")
        End If
        ListBox1.Items.Add(TextBox1.Text)
        TextBox1.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex = -1 Then
            MsgBox("Please select a grade to remove")
            Exit Sub
        End If
        ListBox1.Items.Remove(ListBox1.SelectedItem)
    End Sub
End Class
