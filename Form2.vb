Imports System.Windows.Forms

Imports System.Data.OleDb

Imports System.Globalization

Public Class Form2

Dim con As New OleDbConnection

Dim cmd As New OleDbCommand

Dim bid, cost, noofbooks As Integer

Dim bname, aname, str As String

Dim bdate As Date

Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Book shop\books.mdb")

con.Open()

End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

TextBox1.Text = ""

TextBox2.Text = ""

TextBox3.Text = ""

TextBox4.Text = ""

TextBox5.Text = ""

End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

MDIParent1.Show()

Me.Hide()

End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

bdate = DateTime.Now

str = "insert into stock(bid,bname,aname,cost,noofbooks,bdate) values ( " & bid & ", '" & bname & "' , '" & aname & "' , " & cost & " ,

" & noofbooks & ",'" & bdate & "');"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

MsgBox("Record added")

End Sub

Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox1.TextChanged

bid = Val(TextBox1.Text)

End Sub

Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox2.TextChanged

bname = TextBox2.Text

End Sub

Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox3.TextChanged

aname = TextBox3.Text

End Sub

Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox4.TextChanged

cost = Val(TextBox4.Text)

End Sub

Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox5.TextChanged

noofbooks = Val(TextBox5.Text)

End Sub

End Class
