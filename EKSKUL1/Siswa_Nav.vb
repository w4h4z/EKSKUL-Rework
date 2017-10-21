Imports System.Data.SqlClient
Public Class Siswa_Nav
    Dim db As New Database
    Dim dbcomm As New SqlCommand
    Dim dbread As SqlDataReader
    Dim sql As String
    Dim mid As String
    Dim final As String
    Public idSiswa As String
    Private Sub Siswa_Nav_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        db.conn()
        data()
    End Sub

    Private Sub data()
        sql = "select * from Nilai_Ekskul as ne join siswa as s on ne.id_siswa=s.id_siswa join Ekskul as e on ne.id_ekskul=e.id_ekskul where s.id_siswa='" & idSiswa & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            mid = dbread("soal_mid")
            final = dbread("soal_final")
            dbread.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(mid)
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Process.Start(final)
        Catch ex As Exception
            MsgBox("Soal belum diupload")
        End Try
    End Sub

    Private Sub btnUploadMid_Click(sender As Object, e As EventArgs) Handles btnUploadMid.Click
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "PDF Files(*.pdf)|*.pdf|DOCX Files(*.docx)|*.docx"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
            db.jawabanMid(TextBox1.Text, idSiswa)
            MsgBox("Upload success", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub btnUploadFinal_Click(sender As Object, e As EventArgs) Handles btnUploadFinal.Click
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "PDF Files(*.pdf)|*.pdf|DOCX Files(*.docx)|*.docx"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox2.Text = OpenFileDialog1.FileName
            db.jawabanFinal(TextBox2.Text, idSiswa)
            MsgBox("Upload success", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Login.Show()
        Me.Close()
    End Sub
End Class