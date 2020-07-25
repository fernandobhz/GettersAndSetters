Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim r As String = String.Empty
        Dim c As String = String.Empty

        For Each s As String In Me.TextBox1.Lines
            s = Trim(s)
            Dim x As String() = s.Split(" ")
            Dim var As String = x(1)
            Dim tip As String = x(3)

            r = r & "Function get" & var & "() as " & tip & vbCrLf
            r = r & "   Return Me." & var & vbCrLf
            r = r & "End Function" & vbCrLf

            r = r & vbCrLf

            r = r & "Sub set" & var & "(Byval " & var & " as " & tip & ")" & vbCrLf
            r = r & "   me." & var & " = " & var & vbCrLf
            r = r & "End Sub" & vbCrLf

            r = r & vbCrLf

            c = c & "Me.set" & var & "(" & var & ":=" & var & ")" & vbCrLf
        Next

        Me.TextBox2.Text = r
        Me.TextBox3.Text = c
    End Sub
End Class
