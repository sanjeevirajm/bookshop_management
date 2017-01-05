Imports System.Windows.Forms

Imports System.Data.OleDb

Imports System.Globalization

Public Class Form3

Dim con As New OleDbConnection

Dim cmd As New OleDbCommand

Dim dr As OleDbDataReader

Dim bid, cost, noofbooks, orderno, billno, total, fcost As Integer

Dim bname, aname, toname, str As String

Dim bdate As Date

Dim exc As Boolean

Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

exc = True

con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Book shop\books.mdb")

con.Open()

str = "delete * from cursales"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

str = "select billno from sales"

cmd = New OleDbCommand(str, con)

dr = cmd.ExecuteReader

If dr.HasRows = True Then

Do While dr.Read()

billno = dr.Item("billno")

Loop

End If

If billno < 100 Then

billno = 100

Else

billno = billno + 1

End If

TextBox1.Text = billno

str = "select orderno from cursales"

cmd = New OleDbCommand(str, con)

dr = cmd.ExecuteReader

If dr.HasRows = True Then

Do While dr.Read()

orderno = dr.Item("orderno")

Loop

End If

If orderno < 1 Then

orderno = 1

Else

orderno = orderno + 1

End If

TextBox2.Text = orderno

str = "select bid from stock"

cmd = New OleDbCommand(str, con)

dr = cmd.ExecuteReader

ListBox1.Items.Clear()

If dr.HasRows = True Then

Do While dr.Read()

bid = dr.Item("bid")

ListBox1.Items.Add(bid)

Loop

End If

End Sub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

TextBox1.Text = ""

TextBox2.Text = ""

TextBox3.Text = ""

TextBox4.Text = ""

TextBox5.Text = ""

TextBox6.Text = ""

TextBox7.Text = ""

TextBox8.Text = ""

TextBox9.Text = ""

ListBox1.Items.Clear()

End Sub

Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

MDIParent1.Show()

Me.Hide()

End Sub

Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox1.TextChanged

End Sub

Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox2.TextChanged

orderno = Val(TextBox2.Text)

End Sub

Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox3.TextChanged

bname = TextBox3.Text

End Sub

Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox4.TextChanged

aname = TextBox4.Text

End Sub

Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox5.TextChanged

toname = TextBox5.Text

End Sub

Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox6.TextChanged

bname = TextBox3.Text

End Sub

Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

ListBox1.SelectedIndexChanged

bid = ListBox1.SelectedItem

Dim com As New OleDbCommand

com = New OleDbCommand("select * from stock where bid=" & bid, con)

Dim dr As OleDbDataReader = com.ExecuteReader

If dr.HasRows = True Then

dr.Read()

bname = dr.Item("bname")

aname = dr.Item("aname")

cost = dr.Item("cost")

End If

TextBox3.Text = bname

TextBox4.Text = aname

TextBox6.Text = cost

End Sub

Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox7.TextChanged

noofbooks = Val(TextBox7.Text)

total = cost * noofbooks

TextBox8.Text = total

fcost = fcost + total

TextBox9.Text = fcost

End Sub

Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox8.TextChanged

End Sub

Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles

TextBox9.TextChanged

End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

bdate = DateTime.Now

str = "select orderno from cursales"

cmd = New OleDbCommand(str, con)

dr = cmd.ExecuteReader

If dr.HasRows = True Then

Do While dr.Read()

orderno = dr.Item("orderno")

Loop

End If

If orderno < 1 Then

orderno = 1

Else

orderno = orderno + 1

End If

TextBox2.Text = orderno

str = "insert into cursales values ( " & bid & ", '" & bname & "' , '" & aname & "' , " & cost & " , " & noofbooks & ",'" & bdate & "'," &

billno & "," & orderno & ", '" & toname & "', " & total & ", " & fcost & ");"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

str = "insert into sales values ( " & bid & ", '" & bname & "' , '" & aname & "' , " & cost & " , " & noofbooks & ",'" & bdate & "'," & billno

& "," & orderno & ", '" & toname & "', " & total & ", " & fcost & ");"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

MsgBox("Record added")

TextBox1.Text = billno

TextBox3.Text = ""

TextBox4.Text = ""

TextBox5.Text = ""

TextBox6.Text = ""

TextBox7.Text = ""

TextBox8.Text = ""

TextBox9.Text = ""

End Sub

Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

str = "insert into cursales values ( " & bid & ", '" & bname & "' , '" & aname & "' , " & cost & " , " & noofbooks & ",'" & bdate & "'," &

billno & "," & orderno & ", '" & toname & "', " & total & ", " & fcost & ");"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

str = "insert into sales(bid,bname,aname,cost,noofbooks,bdate,billno,orderno,toname,total,fcost) values ( " & bid & ", '" & bname &

"' , '" & aname & "' , " & cost & " , " & noofbooks & ",'" & bdate & "'," & billno & "," & orderno & ", '" & toname & "', " & total & ", " & fcost &

");"

cmd = New OleDbCommand(str, con)

cmd.ExecuteNonQuery()

MsgBox("Record added")

exc = False

Me.Hide()

MDIParent1.Show()

End Sub

End Class
