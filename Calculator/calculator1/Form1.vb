Public Class frmCalculator

    Private LastOperator As String
    Private FirstPart As Double

    Private PreviousLastOperator As String
    Private PreviousLastPart As Double

    ' If any of the number buttons or the decimal point is clicked

    Private Sub btns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click, btn10.Click, btn11.Click

        ' If the user has just clicked the equals button
        If Label1.Text.EndsWith("= ") Then

            ' Call the clear subroutine
            btnClear.PerformClick()

        End If

        ' Textbox is set equal to the button
        TextBox1.Text &= DirectCast(sender, Button).Text

    End Sub

    ' If any of the plus, subtract, multiply or divide buttons are clicked

    Private Sub btnOperator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles btnOperator1.Click, btnOperator2.Click, btnOperator3.Click, btnOperator4.Click

        Dim NextPart As Double

        Double.TryParse(TextBox1.Text, nextPart)


        If NextPart = 0 Then

            If DirectCast(sender, Button).Text = "-" AndAlso TextBox1.Text = "" Then

                TextBox1.Text &= "-"

            End If

            Return

        End If

        PreviousLastOperator = LastOperator
        PreviousLastPart = NextPart

        ' Calculates the number
        methods.calculate(TextBox1, FirstPart, NextPart, LastOperator)

        ' If there hasn't been a previous operator
        If LastOperator <> "" Then

            Label1.Text &= NextPart.ToString

        Else

            Label1.Text &= FirstPart.ToString

        End If

        LastOperator = DirectCast(sender, Button).Text

        Label1.Text &= " " & LastOperator & " "
        TextBox1.Text = ""

    End Sub

    ' If the equals button is clicked

    Private Sub btnEquals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEquals.Click

        Dim LastPart As Double

        Double.TryParse(TextBox1.Text, LastPart)

        If lastPart <> 0 Then

            ' Do the calculation
            methods.calculate(TextBox1.Text, FirstPart, LastPart, LastOperator)
            Label1.Text &= " " & LastPart.ToString & " = "

        Else

            ' If last part is equal to 0
            methods.changeText(Label1.Text, " = ")
            TextBox1.Text = FirstPart.ToString

        End If

        ' Reset
        FirstPart = 0
        LastOperator = ""

    End Sub

    ' If the 'C' button is clicked

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click

        ' Resets the calculator

        TextBox1.Text = ""
        FirstPart = 0
        LastOperator = ""
        Label1.Text = ""
        PreviousLastOperator = ""
        PreviousLastPart = 0

    End Sub

    ' If a key from the keyboard is pressed

    Private Sub Form1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress

        ' Allows the user to use the calculator using their keyboard

        Dim ktp = From btn As Button In Me.Controls.OfType(Of Button)() _
                  Where btn.Text.ToLower = Char.ToLower(e.KeyChar) _
                  Select btn

        If ktp.Count > 0 Then

            ktp.First.PerformClick()

        End If

    End Sub

End Class