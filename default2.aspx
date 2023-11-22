Imports excel = Microsoft.Office.Interop.Excel
‘untuk koneksi ke Ms. Access
Imports System.Data.OleDb
Partial Class Default2
Inherits System.Web.UI.Page
‘untuk koneksi ke Ms. Access
Dim koneksi As New OleDbConnection(“Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Eresha 2013\dotNET\ASP\SAB.accdb”)
Sub kosong()
TextBox6.Text = “”
TextBox7.Text = “”
TextBox8.Text = “”
TextBox9.Text = “”
TextBox10.Text = “”
End Sub
Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
Dim ObjAppExcel As New excel.Application
Dim ObjDocExcel = ObjAppExcel.Workbooks.Open(“D:\Eresha 2013\dotNET\ASP\SAB.xls”)
Dim urutan As Integer
urutan = ObjAppExcel.Range(“F1”).Value
urutan = urutan + 1
ObjAppExcel.Range(“A” & urutan).Insert()
ObjAppExcel.Range(“A” & urutan).Value = TextBox6.Text
ObjAppExcel.Range(“B” & urutan).Insert()
ObjAppExcel.Range(“B” & urutan).Value = TextBox7.Text
ObjAppExcel.Range(“C” & urutan).Insert()
ObjAppExcel.Range(“C” & urutan).Value = TextBox8.Text
ObjAppExcel.Range(“D” & urutan).Insert()
ObjAppExcel.Range(“D” & urutan).Value = TextBox9.Text
ObjAppExcel.Range(“E” & urutan).Insert()
ObjAppExcel.Range(“E” & urutan).Value = TextBox10.Text
ObjAppExcel.Range(“F1”).Value = urutan
ObjDocExcel.Save()
ObjDocExcel.Close()
ObjAppExcel.Quit()
Call kosong()
End Sub

Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
Call kosong()
Button1.Enabled = False
Button2.Enabled = True
Button3.Enabled = False
TextBox6.Visible = False
TextBox7.Visible = False
TextBox8.Visible = False
TextBox9.Visible = False
TextBox10.Visible = False
End Sub

Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
If Button2.Text = “INSERT” Then
Button1.Enabled = True
Button3.Enabled = True
Button2.Text = “CANCEL”
TextBox6.Focus()
ElseIf Button2.Text = “CANCEL” Then
Button1.Enabled = False
Button3.Enabled = False
Button2.Text = “INSERT”
End If
End Sub
‘untuk simpan ke MS. Access
Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
koneksi.Close()
koneksi.Open()
Dim simpan As New OleDbCommand
simpan.Connection = koneksi
simpan.CommandType = Data.CommandType.Text
simpan.CommandText = “INSERT INTO SAB (fName,fNickName,fAddress,fPhone,fEmail) VALUES (‘” & TextBox11.Text & “‘,'” & TextBox12.Text & “‘,'” & TextBox13.Text & “‘,'” & TextBox14.Text & “‘,'” & TextBox15.Text & “‘)”
simpan.ExecuteNonQuery()
koneksi.Close()
MsgBox(“Data Tersimpan”)
GridView1.DataBind()
End Sub
End Class
