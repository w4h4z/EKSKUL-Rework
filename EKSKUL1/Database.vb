Imports System.Data.SqlClient
Public Class Database
    Dim dbconn As New SqlConnection
    Dim dbcomm As New SqlCommand
    Dim sql As String
    Dim dbread As SqlDataReader
    Dim lastId As String
    Public Function conn()
        dbconn = New SqlConnection("data source=.\SQLEXPRESS;database=EKSKUL1;integrated security=true")

        Try
            dbconn.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dbconn
    End Function

    Public Function crud(sql)
        Try
            dbcomm = New SqlCommand(sql, conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            Return dbread
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Function

    Public Function crudid(sql)
        Try
            dbcomm = New SqlCommand(sql, conn)
            lastId = dbcomm.ExecuteScalar
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return lastId
    End Function

#Region "Guru"
    Public Sub insertGuru(id, nama, foto)
        sql = "insert into Guru(id_guru,nama_guru,foto_guru) values('" & id & "','" & nama & "','" & foto & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteGuru(id)
        sql = "delete Guru where id_guru='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateGuru(id, nama, foto)
        sql = "update Guru set nama_guru='" & nama & "',foto_guru='" & foto & "' where id_guru='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Siswa"
    Public Sub insertSiswa(id, nama, tgl, telp, jk, asal, foto)
        sql = "insert into Siswa(id_siswa,nama_siswa,tgl_lhr_siswa,telp_siswa,jenis_kelamin,asal_sekolah,foto_siswa) values('" & id & "','" & nama & "','" & tgl & "','" & telp & "','" & jk & "','" & asal & "','" & foto & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateSiswa(id, nama, tgl, telp, jk, asal, foto)
        sql = "update Siswa set nama_siswa='" & nama & "', tgl_lhr_siswa='" & tgl & "',telp_siswa='" & telp & "',jenis_kelamin='" & jk & "',asal_sekolah='" & asal & "',foto_siswa='" & foto & "' where id_siswa='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteSiswa(id)
        sql = "delete Siswa where id_siswa='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Ekskul"
    Public Sub insertEkskul(id, guru, nama, foto)
        sql = "insert into Ekskul(id_ekskul,id_guru,nama_ekskul,foto_ekskul) values('" & id & "','" & guru & "','" & nama & "','" & foto & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteEkskul(id)
        sql = "delete Ekskul where id_ekskul='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateEkskul(id, guru, nama, foto)
        sql = "update Ekskul set id_guru='" & guru & "',nama_ekskul='" & nama & "',foto_ekskul='" & foto & "' where id_ekskul='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Jadwal"
    Public Sub insertJadwal(id, ekskul, hari, jam)
        sql = "insert into Jadwal_Ekskul(id_jadwal,id_ekskul,hari_ekskul,jam_ekskul) values('" & id & "','" & ekskul & "','" & hari & "','" & jam & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateJadwal(id, ekskul, hari, jam)
        sql = "update Jadwal_Ekskul set id_ekskul='" & ekskul & "',hari_ekskul='" & hari & "',jam_ekskul='" & jam & "' where id_jadwal='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteJadwal(id)
        sql = "delete Jadwal_Ekskul where id_jadwal='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Alokasi siswa ekskul"
    Public Sub insertToEkskul(siswa, ekskul)
        sql = "update Siswa set id_ekskul='" & ekskul & "' where id_siswa='" & siswa & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteFromEkskul(siswa)
        sql = "update Siswa set id_ekskul=NULL where id_siswa='" & siswa & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Nilai ekskul"
    Public Sub insertSoalMid(nilai, siswa, ekskul, mid)
        sql = "insert into Nilai_Ekskul(id_nilai,id_siswa,id_ekskul,soal_mid) values('" & nilai & "','" & siswa & "','" & ekskul & "','" & mid & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub insertSoalFinal(nilai, siswa, ekskul, mid, final)
        sql = "insert into Nilai_Ekskul(id_nilai,id_siswa,id_ekskul,soal_mid,soal_final) values('" & nilai & "','" & siswa & "','" & ekskul & "','" & mid & "','" & final & "')"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateSoalMid(nilai, siswa, ekskul, mid)
        sql = "update Nilai_Ekskul set id_siswa='" & siswa & "',id_ekskul='" & ekskul & "',soal_mid='" & mid & "' where id_nilai='" & nilai & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateSoalFinal(nilai, siswa, ekskul, mid, final)
        sql = "update Nilai_Ekskul set id_siswa='" & siswa & "',id_ekskul='" & ekskul & "',soal_mid='" & mid & "',soal_final='" & final & "' where id_nilai='" & nilai & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateNilaiMid(mid, id)
        sql = "update Nilai_Ekskul set nilai_mid='" & mid & "' where id_nilai='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub updateNilaiFinal(final, id)
        sql = "update Nilai_Ekskul set nilai_final='" & final & "' where id_nilai='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub deleteNilai(id)
        sql = "delete Nilai_Ekskul where id_nilai='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region

#Region "Siswa Nav"
    Public Sub jawabanMid(mid, id)
        sql = "update Nilai_Ekskul set jawaban_mid='" & mid & "' where id_siswa='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub

    Public Sub jawabanFinal(final, id)
        sql = "update Nilai_Ekskul set jawaban_final='" & final & "' where id_siswa='" & id & "'"

        crud(sql)
        dbread.Close()
    End Sub
#End Region
End Class
