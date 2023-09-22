Imports System.Math
Class Form2

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim p, p1, p2, d, z, pval As Single
        Dim R As Integer
        p1 = TextBox1.Text
        p2 = TextBox2.Text
        R = TextBox3.Text

        p = (p1 + p2) / 2
        d = Sqrt(2 * p * (1 - p) / R)

        z = (p1 - p2) / d
        pval = 1 - ztest(z)
        If pval < TextBox4.Text Then
            Label5.Text = "Result: the bootstrap is significantly higher, pval=" & pval
        Else
            Label5.Text = "Result: the bootstrap is NOT significantly higher, pval=" & pval
        End If
    End Sub
    Function fact(ByVal x As Double)
        Dim a As Integer = x
        Do While a > 0
            x = x * (x - 1)
            a = a - 1
        Loop
        Return x
    End Function
    Function ztest(ByVal x As Single) As Single

        Dim phi, t As Double


        Dim p As Double = 0.2316419
        Dim b1 As Double = 0.31938153
        Dim b2 As Double = -0.356563782
        Dim b3 As Double = 1.781477937
        Dim b4 As Double = -1.821255978
        Dim b5 As Double = 1.330274429

        Dim z As Double = Math.Exp(-0.5 * Math.Pow(x, 2)) / Sqrt(2 * PI)
        t = 1 / (1 + p * x)

        phi = 1 - z * (b1 * t + (b2 * t ^ 2) + (b3 * t ^ 3) + (b4 * t ^ 4) + (b5 * t ^ 5)) + (7.5 * 10 ^ -8)
        Return phi
    End Function
    
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox2.KeyPress, TextBox3.KeyPress, TextBox4.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "."c Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class