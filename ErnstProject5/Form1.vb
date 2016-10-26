Public Class Form1
    ' init vars
    Dim Average, Total, H, AWH, R As Decimal

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'reset vars and clear all boxes and user inputs
        H = 0
        Average = 0
        Total = 0
        AWH = 0
        ListBox1.Items.Clear()
        TextBox1.Text = ""
        Label1.Text = "Average"
        Label2.Text = "Average without outlier"
        Button3.Visible = True

    End Sub

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
        If Not IsNumeric(TextBox2.Text) Then
            TextBox1.Clear()
            MsgBox("Please only enter numbers")
            Exit Sub
        End If
        If Val(TextBox2.Text) <= 0 Then
            TextBox1.Clear()
            MsgBox("Please enter positave values")
            Exit Sub
        End If
        R = Val(TextBox2.Text)
        If ListBox1.Items.Count = 10 Then
            TextBox1.Clear()
            MsgBox("No more bills can be entered")
        End If
        'end error check
        'math for bills
        If Val(TextBox1.Text) <= 1000 Then
            TextBox1.Text = ((Val(TextBox1.Text) * R) + 12)
        End If
        If Val(TextBox1.Text) > 1000 Then
            TextBox1.Text = ((Val(TextBox1.Text) * R) + 20)
        End If
        ListBox1.Items.Add(TextBox1.Text)
        TextBox1.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex = -1 Then
            MsgBox("Please select a bill to remove")
            Exit Sub
        End If
        ListBox1.Items.Remove(ListBox1.SelectedItem)
        'remove selected bill
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'error check
        If ListBox1.Items.Count = 0 Or ListBox1.Items.Count = 1 Then
            MsgBox("Please enter 2 or more bills ")
            Exit Sub
        End If
        'find average with or without heighest value
        For counter = 0 To ListBox1.Items.Count - 1
            Total = Total + Val(ListBox1.Items(counter))
            If Val(ListBox1.Items(counter)) > H Then
                H = (ListBox1.Items(counter))
            End If
        Next
        Average = Total / ListBox1.Items.Count
        AWH = ((Total - H) / (ListBox1.Items.Count - 1))
        Label1.Text = FormatCurrency(Average)
        Label2.Text = FormatCurrency(AWH)
        Button3.Visible = False
    End Sub
End Class
