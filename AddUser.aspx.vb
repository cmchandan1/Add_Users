Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Security.Cryptography
Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Partial Class AddUser
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then
            Dim constr As OleDbConnection
            Dim path = System.IO.Path.GetFullPath(Server.MapPath("~/database.xlsx"))
            constr = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';")
            constr.Open()
            Dim oleda As OleDbDataAdapter = New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            Using cmd As OleDbCommand = New OleDbCommand("SELECT * FROM [Users$]")
                Using sda As New SqlDataAdapter()
                    Dim dt As New DataTable()
                    cmd.CommandType = CommandType.Text
                    cmd.Connection = constr
                    oleda.SelectCommand = cmd
                    oleda.Fill(dt)
                    gvUsers.DataSource = dt
                    gvUsers.DataBind()
                    constr.Close()
                End Using
            End Using
        End If
    End Sub

    Protected Sub OnRowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(2).Text = Decrypt(e.Row.Cells(2).Text)
        End If
    End Sub
    Protected Sub Submit(sender As Object, e As EventArgs)
        Dim constr1 As OleDbConnection
        Dim path = System.IO.Path.GetFullPath(Server.MapPath("~/database.xlsx"))
        constr1 = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;YES;';")
        Dim Name As String = txtUsername.Text.Trim()
        Dim Pass As String = Encrypt(txtPassword.Text.Trim())


        Using cmd As OleDbCommand = New OleDbCommand()
            cmd.Connection = constr1
            constr1.Open()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "Insert into [Users$] ([Uname], [MyPass]) VALUES ('" & Name & "', '" & Pass & "')"
            cmd.ExecuteNonQuery()
            constr1.Close()
        End Using
        Response.Redirect(Request.Url.AbsoluteUri)
    End Sub

    Private Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Private Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function
End Class
