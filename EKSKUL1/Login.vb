Imports System.Data.SqlClient
Public Class Login
    Dim dbcomm As New SqlCommand
    Dim dbread As SqlDataReader
    Dim sql As String
    Dim db As New Database

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        db.conn
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim b As Boolean
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("All data must be fill!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If TextBox1.Text.Substring(0, 1) = "S" And TextBox2.Text = "siswa" Then
            sql = "select * from Siswa where id_siswa='" & TextBox1.Text & "'"

            Try
                dbcomm = New SqlCommand(sql, db.conn)
                dbread = dbcomm.ExecuteReader
                If dbread.HasRows Then
                    b = True
                    Siswa_Nav.idSiswa = TextBox1.Text
                    Siswa_Nav.Show()
                    Me.Close()
                Else
                    MsgBox("Login failed, check your username and password!", MsgBoxStyle.Exclamation)
                End If
                dbread.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try

        ElseIf TextBox1.Text.Substring(0, 1) = "G" And TextBox2.Text = "guru" Then
            sql = "select * from Guru where id_guru='" & TextBox1.Text & "'"

            Try
                dbcomm = New SqlCommand(sql, db.conn)
                dbread = dbcomm.ExecuteReader
                If dbread.HasRows Then
                    b = True
                    Admin_Nav.Show()
                    Me.Close()
                Else
                    MsgBox("Login failed, check your username and password!", MsgBoxStyle.Exclamation)
                End If
                dbread.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        End If

        If b = False Then
            MsgBox("Login failed, check your username and password!", MsgBoxStyle.Exclamation)
        End If
    End Sub
End Class