Public Class Form1

Dim pwd As String

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

If TextBox1.Text = "Admin" And TextBox2.Text = pwd Then

MsgBox("Login successful")

MDIParent1.Show()

Me.Hide()

Else

MsgBox("Incorrect username or password")

TextBox2.Text = ""

End If

End Sub

Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

TextBox1.Text = "Admin"

pwd = "book15"

End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

TextBox2.Text = ""

End Sub

End Class
