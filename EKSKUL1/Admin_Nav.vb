Imports System.Data.SqlClient
Public Class Admin_Nav
    Dim db As New Database
    Dim dbcomm As New SqlCommand
    Dim sql As String
    Dim dbread As SqlDataReader
    Dim locationFoto As String
    Private Sub Admin_Nav_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'EKSKUL1DataSet.Ekskul' table. You can move, or remove it, as needed.
        Me.EkskulTableAdapter.Fill(Me.EKSKUL1DataSet.Ekskul)
        'TODO: This line of code loads data into the 'EKSKUL1DataSet.Guru' table. You can move, or remove it, as needed.
        Me.GuruTableAdapter.Fill(Me.EKSKUL1DataSet.Guru)
        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed

        db.conn()
        dataGuru()
        dataSiswa()
        dataEkskul()
        dataJadwal()
        maxKodeGuru()
        maxKodeSiswa()
        maxKodeEkskul()
        maxKodeJadwal()
        siswaNull()
        siswaNotNull()
        siswaEkskul()
        dataNilai()
        maxKodeNilai()
    End Sub

    Private Function autoIncrement(kode)
        Dim kodeId As String
        Dim huruf As String = kode.substring(0, 1)
        Dim id As String = kode.substring(1)
        Dim angka As Integer = Integer.Parse(id)
        angka += 1
        kodeId = huruf & angka.ToString("D" & id.Length)
        Return kodeId
    End Function

#Region "Side Bar"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectTab(0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectTab(1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectTab(2)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabControl1.SelectTab(3)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TabControl1.SelectTab(4)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab(5)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TabControl1.SelectTab(6)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Login.Show()
        Me.Close()
    End Sub

#End Region

#Region "Guru"
    Private Sub dataGuru()
        DataGridViewGuru.Rows.Clear()
        sql = "select * from Guru"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            Dim i As Integer = 1
            While dbread.Read
                Dim foto As String = dbread("foto_guru")
                DataGridViewGuru.Rows.Add(i, dbread("id_guru"), dbread("nama_guru"), foto.Substring(foto.LastIndexOf("\") + 1), dbread("foto_guru"))
                i += 1
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub maxKodeGuru()
        sql = "select max(id_guru) as lastId from Guru"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            txtIdGuru.Text = autoIncrement(dbread("lastId"))
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub resetGuru()
        txtNamaGuru.Text = ""
        PictureBoxGuru.Image = My.Resources.profil
        locationFoto = ""
        maxKodeGuru()
    End Sub

    Private Sub btnUploadGuru_Click(sender As Object, e As EventArgs) Handles btnUploadGuru.Click
        If OpenFileDialog1.FileName = "" Then
            MsgBox("Empty!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locationFoto = OpenFileDialog1.FileName
            PictureBoxGuru.ImageLocation = locationFoto
        End If
    End Sub

    Private Sub btnAddGuru_Click(sender As Object, e As EventArgs) Handles btnAddGuru.Click
        If txtNamaGuru.Text = "" Or locationFoto = "" Then
            MsgBox("All data must be fill!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        db.insertGuru(txtIdGuru.Text, txtNamaGuru.Text, locationFoto)
        dataGuru()
        resetGuru()
        txtIdGuru.Text = autoIncrement(txtIdGuru.Text)
        MsgBox("Insert data succcess", MsgBoxStyle.Information)
    End Sub

    Private Sub btnEditGuru_Click(sender As Object, e As EventArgs) Handles btnEditGuru.Click
        If txtNamaGuru.Text = "" Or locationFoto = "" Then
            MsgBox("All data must be fill!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Try
            db.updateGuru(txtIdGuru.Text, txtNamaGuru.Text, locationFoto)
            MsgBox("update data success", MsgBoxStyle.Information)
            dataGuru()
            resetGuru()
            maxKodeGuru()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridViewGuru_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewGuru.CellClick
        If e.ColumnIndex = 5 Then
            Dim i As Integer = DataGridViewGuru.CurrentRow.Index
            db.deleteGuru(DataGridViewGuru.Item(1, i).Value)
            MsgBox("Delete data success", MsgBoxStyle.Information)
            dataGuru()
            resetGuru()
            maxKodeGuru()
        End If
    End Sub

    Private Sub btnCancelGuru_Click(sender As Object, e As EventArgs) Handles btnCancelGuru.Click
        resetGuru()
    End Sub

    Private Sub DataGridViewGuru_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewGuru.CellMouseClick
        Dim i As Integer = DataGridViewGuru.CurrentRow.Index

        txtIdGuru.Text = DataGridViewGuru.Item(1, i).Value
        txtNamaGuru.Text = DataGridViewGuru.Item(2, i).Value
        PictureBoxGuru.ImageLocation = DataGridViewGuru.Item(4, i).Value '
        locationFoto = DataGridViewGuru.Item(4, i).Value '
    End Sub
#End Region

#Region "Siswa"
    Public Sub dataSiswa()
        DataGridViewSiswa.Rows.Clear()
        sql = "select * from Siswa"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            Dim no As Integer = 1
            While dbread.Read
                Dim foto As String = dbread("foto_siswa")
                DataGridViewSiswa.Rows.Add(no, dbread("id_siswa"), dbread("nama_siswa"), dbread("tgl_lhr_siswa").toshortdatestring, dbread("telp_siswa"), dbread("jenis_kelamin"), dbread("asal_sekolah"), foto.Substring(foto.LastIndexOf("\") + 1), foto)
                no += 1
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub maxKodeSiswa()
        sql = "select max(id_siswa) as lastId from Siswa"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            txtIdSiswa.Text = autoIncrement(dbread("lastId"))
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub btnUploadSiswa_Click(sender As Object, e As EventArgs) Handles btnUploadSiswa.Click
        If OpenFileDialog1.FileName = "" Then
            MsgBox("Empty!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locationFoto = OpenFileDialog1.FileName
            PictureBoxSiswa.ImageLocation = locationFoto
        End If
    End Sub

    Private Sub btnAddSiswa_Click(sender As Object, e As EventArgs) Handles btnAddSiswa.Click
        Dim jk As String
        If rbLkSiswa.Checked = True Then
            jk = "Laki-Laki"
        ElseIf rbPrSiswa.Checked = True
            jk = "Perempuan"
        End If
        db.insertSiswa(txtIdSiswa.Text, txtNamaSiswa.Text, DateTimePickerSiswa.Value.ToString("yyyy/MM/dd"), txtTelpSiswa.Text, jk, txtAsalSiswa.Text, locationFoto)
        MsgBox("Insert data success", MsgBoxStyle.Information)
        dataSiswa()
        resetSiswa()
        maxKodeSiswa()
    End Sub

    Private Sub btnEditSiswa_Click(sender As Object, e As EventArgs) Handles btnEditSiswa.Click
        Dim jk As String
        If rbLkSiswa.Checked = True Then
            jk = "Laki-Laki"
        ElseIf rbPrSiswa.Checked = True
            jk = "Perempuan"
        End If
        db.updateSiswa(txtIdSiswa.Text, txtNamaSiswa.Text, DateTimePickerSiswa.Value.ToString("yyyy/MM/dd"), txtTelpSiswa.Text, jk, txtAsalSiswa.Text, locationFoto)
        MsgBox("update data success", MsgBoxStyle.Information)
        dataSiswa()
        resetSiswa()
        maxKodeSiswa()
    End Sub

    Private Sub btnDeleteSiswa_Click(sender As Object, e As EventArgs) Handles btnDeleteSiswa.Click
        db.deleteSiswa(txtIdSiswa.Text)
        MsgBox("Delete data success", MsgBoxStyle.Information)
        dataSiswa()
        resetSiswa()
        maxKodeSiswa()
    End Sub

    Private Sub DataGridViewSiswa_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewSiswa.CellMouseClick
        Dim i As Integer = DataGridViewSiswa.CurrentRow.Index

        txtIdSiswa.Text = DataGridViewSiswa.Item(1, i).Value
        txtNamaSiswa.Text = DataGridViewSiswa.Item(2, i).Value
        DateTimePickerSiswa.Value = DataGridViewSiswa.Item(3, i).Value
        txtTelpSiswa.Text = DataGridViewSiswa.Item(4, i).Value
        If DataGridViewSiswa.Item(5, i).Value = "Laki-Laki" Then
            rbLkSiswa.Checked = True
        ElseIf DataGridViewSiswa.Item(5, i).Value = "Perempuan" Then
            rbPrSiswa.Checked = True
        End If
        txtAsalSiswa.Text = DataGridViewSiswa.Item(6, i).Value
        locationFoto = DataGridViewSiswa.Item(8, i).Value
        PictureBoxSiswa.ImageLocation = DataGridViewSiswa.Item(8, i).Value
    End Sub

    Private Sub resetSiswa()
        txtIdSiswa.Text = ""
        txtNamaSiswa.Text = ""
        DateTimePickerSiswa.Value = Date.Now
        txtTelpSiswa.Text = ""
        rbLkSiswa.Checked = False
        rbPrSiswa.Checked = False
        txtAsalSiswa.Text = ""
        locationFoto = ""
        PictureBoxSiswa.Image = My.Resources.profil
        maxKodeSiswa()
    End Sub

    Private Sub btnCancelSiswa_Click(sender As Object, e As EventArgs) Handles btnCancelSiswa.Click
        resetSiswa()
    End Sub
#End Region

#Region "Ekskul"
    Public Sub dataEkskul()
        DataGridViewEkskul.Rows.Clear()
        sql = "select * from Ekskul as e join Guru as g on e.id_guru=g.id_guru"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            Dim no = 1
            While dbread.Read
                Dim foto As String = dbread("foto_guru")
                DataGridViewEkskul.Rows.Add(no, dbread("id_ekskul"), dbread("nama_ekskul"), dbread("id_guru"), dbread("nama_guru"), foto.Substring(foto.LastIndexOf("\") + 1), foto)
                no += 1
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub maxKodeEkskul()
        sql = "select max(id_ekskul) as lastId from Ekskul"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            txtIdEkskul.Text = autoIncrement(dbread("lastId"))
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub resetEkskul()
        txtNamaEkskul.Text = ""
        PictureBoxEkskul.Image = My.Resources.profil
        cbGuruEkskul.SelectedIndex = 0
        maxKodeEkskul()
    End Sub

    Private Sub btnUploadEkskul_Click(sender As Object, e As EventArgs) Handles btnUploadEkskul.Click
        If OpenFileDialog1.FileName = "" Then
            MsgBox("Empty!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locationFoto = OpenFileDialog1.FileName
            PictureBoxEkskul.ImageLocation = locationFoto
        End If
    End Sub

    Private Sub btnAddEkskul_Click(sender As Object, e As EventArgs) Handles btnAddEkskul.Click
        For i As Integer = 0 To DataGridViewEkskul.RowCount - 1
            If DataGridViewEkskul.Item(3, i).Value = cbGuruEkskul.SelectedValue Then
                MsgBox("Guru pembimbing tidak boleh sama!", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        Next
        db.insertEkskul(txtIdEkskul.Text, cbGuruEkskul.SelectedValue, txtNamaEkskul.Text, locationFoto)
        dataEkskul()
        resetEkskul()
        MsgBox("Insert data success", MsgBoxStyle.Information)
    End Sub

    Private Sub btnDeleteEkskul_Click(sender As Object, e As EventArgs) Handles btnDeleteEkskul.Click
        db.deleteEkskul(txtIdEkskul.Text)
        dataEkskul()
        resetEkskul()
        MsgBox("Delete data success", MsgBoxStyle.Information)
    End Sub

    Private Sub DataGridViewEkskul_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewEkskul.CellMouseClick
        Dim i As Integer = DataGridViewEkskul.CurrentRow.Index

        txtIdEkskul.Text = DataGridViewEkskul.Item(1, i).Value
        txtNamaEkskul.Text = DataGridViewEkskul.Item(2, i).Value
        cbGuruEkskul.SelectedValue = DataGridViewEkskul.Item(3, i).Value
        locationFoto = DataGridViewEkskul.Item(6, i).Value
        PictureBoxEkskul.ImageLocation = DataGridViewEkskul.Item(6, i).Value
    End Sub

    Private Sub btnEditEkskul_Click(sender As Object, e As EventArgs) Handles btnEditEkskul.Click
        db.updateEkskul(txtIdEkskul.Text, cbGuruEkskul.SelectedValue, txtNamaEkskul.Text, locationFoto)
        dataEkskul()
        resetEkskul()
        MsgBox("Update data success", MsgBoxStyle.Information)
    End Sub

#End Region

#Region "Jadwal"
    Private Sub dataJadwal()
        DataGridViewJadwal.Rows.Clear()
        sql = "select * from Jadwal_Ekskul as je join Ekskul as e on je.id_ekskul=e.id_ekskul join Guru as g on e.id_guru=g.id_guru"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            Dim no As Integer = 1
            While dbread.Read
                DataGridViewJadwal.Rows.Add(no, dbread("id_jadwal"), dbread("id_ekskul"), dbread("nama_ekskul"), dbread("hari_ekskul"), dbread("jam_ekskul"), dbread("nama_guru"))
                no += 1
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub maxKodeJadwal()
        sql = "select max(id_jadwal) as lastId from Jadwal_Ekskul"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            txtIdJadwal.Text = autoIncrement(dbread("lastId"))
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub resetJadwal()
        cbEkskulJadwal.SelectedIndex = 0
        cbHariJadwal.SelectedIndex = -1
        cbHariJadwal.Text = "Pilih"
        txtJamJadwal.Text = ""
        maxKodeJadwal()
    End Sub

    Private Sub btnAddJadwal_Click(sender As Object, e As EventArgs) Handles btnAddJadwal.Click
        db.insertJadwal(txtIdJadwal.Text, cbEkskulJadwal.SelectedValue, cbHariJadwal.Text, txtJamJadwal.Text)
        dataJadwal()
        resetJadwal()
        MsgBox("Insert data success", MsgBoxStyle.Information)
    End Sub

    Private Sub btnCancelJadwal_Click(sender As Object, e As EventArgs) Handles btnCancelJadwal.Click
        resetJadwal()
    End Sub

    Private Sub DataGridViewJadwal_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewJadwal.CellMouseClick
        Dim i As Integer = DataGridViewJadwal.CurrentRow.Index

        txtIdJadwal.Text = DataGridViewJadwal.Item(1, i).Value
        cbEkskulJadwal.SelectedValue = DataGridViewJadwal.Item(2, i).Value
        cbHariJadwal.Text = DataGridViewJadwal.Item(4, i).Value
        txtJamJadwal.Text = DataGridViewJadwal.Item(5, i).Value
    End Sub

    Private Sub btnDeleteJadwal_Click(sender As Object, e As EventArgs) Handles btnDeleteJadwal.Click
        db.deleteJadwal(txtIdJadwal.Text)
        dataJadwal()
        resetJadwal()
        MsgBox("Delete data success", MsgBoxStyle.Information)
    End Sub

    Private Sub btnEditJadwal_Click(sender As Object, e As EventArgs) Handles btnEditJadwal.Click
        db.updateJadwal(txtIdJadwal.Text, cbEkskulJadwal.SelectedValue, cbHariJadwal.Text, txtJamJadwal.Text)
        dataJadwal()
        resetJadwal()
        MsgBox("Update data success", MsgBoxStyle.Information)
    End Sub

#End Region

#Region "Alokasi siswa ekskul"
    Private Sub siswaNull()
        DataGridViewNull.Rows.Clear()
        sql = "select * from Siswa where id_ekskul is NULL"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewNull.Rows.Add(dbread("id_siswa"), dbread("nama_siswa"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub siswaNotNull()
        DataGridViewNotNull.Rows.Clear()
        sql = "select * from Siswa where id_ekskul='" & cbEkskulAlokasi.SelectedValue & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                DataGridViewNotNull.Rows.Add(dbread("id_siswa"), dbread("nama_siswa"))
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnDeleteFromEkskul_Click(sender As Object, e As EventArgs) Handles btnDeleteFromEkskul.Click
        Dim i As Integer = DataGridViewNotNull.CurrentRow.Index
        db.deleteFromEkskul(DataGridViewNotNull.Item(0, i).Value)
        siswaNull()
        siswaNotNull()
        MsgBox("Delete data success", MsgBoxStyle.Information)
    End Sub

    Private Sub btnInsertToEkskul_Click(sender As Object, e As EventArgs) Handles btnInsertToEkskul.Click
        Dim i As Integer = DataGridViewNull.CurrentRow.Index
        db.insertToEkskul(DataGridViewNull.Item(0, i).Value, cbEkskulAlokasi.SelectedValue)
        siswaNull()
        siswaNotNull()
        MsgBox("Insert data success", MsgBoxStyle.Information)
    End Sub

    Private Sub cbEkskulAlokasi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEkskulAlokasi.SelectedIndexChanged
        siswaNotNull()
    End Sub

#End Region

#Region "Nilai ekskul"
    Private Sub btnMidNilai_Click(sender As Object, e As EventArgs) Handles btnMidNilai.Click
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "PDF Files(*.pdf)|*.pdf|DOCX Files(*.docx)|*.docx"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locationFoto = OpenFileDialog1.FileName
            txtMidNilai.Text = locationFoto
        End If
    End Sub

    Private Sub btnFinalNilai_Click(sender As Object, e As EventArgs) Handles btnFinalNilai.Click
        OpenFileDialog1.Reset()
        OpenFileDialog1.Filter = "PDF Files(*.pdf)|*.pdf|DOCX Files(*.docx)|*.docx"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            locationFoto = OpenFileDialog1.FileName
            txtFinalNilai.Text = locationFoto
        End If
    End Sub

    Private Sub siswaEkskul()
        cbSiswaNilai.Items.Clear()
        sql = "select * from Siswa where id_ekskul='" & cbEkskulNilai.SelectedValue & "'"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            While dbread.Read
                cbSiswaNilai.DisplayMember = "Text"
                cbSiswaNilai.ValueMember = "Value"
                cbSiswaNilai.Items.Add(New With {Key .Text = dbread("nama_siswa"), Key .Value = dbread("id_siswa")})
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub cbEkskulNilai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEkskulNilai.SelectedIndexChanged
        siswaEkskul()
    End Sub

    Private Sub dataNilai()
        DataGridViewNilai.Rows.Clear()
        sql = "select * from Nilai_Ekskul as ne join siswa as s on ne.id_siswa=s.id_siswa join Ekskul as e on ne.id_ekskul=e.id_ekskul"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            Dim no As Integer = 1
            While dbread.Read
                DataGridViewNilai.Rows.Add(no, dbread("id_nilai"), dbread("id_siswa"), dbread("nama_siswa"), dbread("id_ekskul"), dbread("nama_ekskul"), dbread("nilai_mid"), dbread("nilai_final"), dbread("soal_mid"), dbread("soal_final"), dbread("jawaban_mid"), dbread("jawaban_final"))
                no += 1
            End While
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub maxKodeNilai()
        sql = "select max(id_nilai) as lastId from Nilai_Ekskul"

        Try
            dbcomm = New SqlCommand(sql, db.conn)
            dbread = dbcomm.ExecuteReader
            dbread.Read()
            txtIdNilai.Text = autoIncrement(dbread("lastId"))
            dbread.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub resetNilai()
        cbEkskulNilai.SelectedItem = 0
        cbSiswaNilai.SelectedIndex = -1
        cbSiswaNilai.Text = "Pilih"
        txtMidNilai.Text = ""
        txtFinalNilai.Text = ""
        maxKodeNilai()
    End Sub

    Private Sub btnAddNilai_Click(sender As Object, e As EventArgs) Handles btnAddNilai.Click
        If Not txtMidNilai.Text = "" And txtFinalNilai.Text = "" Then
            db.insertSoalMid(txtIdNilai.Text, cbSiswaNilai.SelectedItem.value, cbEkskulNilai.SelectedValue, txtMidNilai.Text)
            dataNilai()
            resetNilai()
            MsgBox("Insert data success", MsgBoxStyle.Information)
            Exit Sub
        End If
        If Not txtFinalNilai.Text = "" And Not txtMidNilai.Text = "" Then
            db.insertSoalFinal(txtIdNilai.Text, cbSiswaNilai.SelectedItem.value, cbEkskulNilai.SelectedValue, txtMidNilai.Text, txtFinalNilai.Text)
            dataNilai()
            resetNilai()
            MsgBox("Insert data success", MsgBoxStyle.Information)
            Exit Sub
        End If
    End Sub

    Private Sub btnCancelNilai_Click(sender As Object, e As EventArgs) Handles btnCancelNilai.Click
        resetNilai()
    End Sub

    Private Sub btnDeleteNilai_Click(sender As Object, e As EventArgs) Handles btnDeleteNilai.Click
        db.deleteNilai(txtIdNilai.Text)
        dataNilai()
        resetNilai()
        MsgBox("Delete data success", MsgBoxStyle.Information)
    End Sub

    Private Sub btnEditNilai_Click(sender As Object, e As EventArgs) Handles btnEditNilai.Click
        If Not txtMidNilai.Text = "" And txtFinalNilai.Text = "" Then
            db.updateSoalMid(txtIdNilai.Text, cbSiswaNilai.SelectedItem.Value, cbEkskulNilai.SelectedValue, txtMidNilai.Text)
            dataNilai()
            resetNilai()
            MsgBox("Update data success", MsgBoxStyle.Information)
            Exit Sub
        End If
        If Not txtFinalNilai.Text = "" And Not txtMidNilai.Text = "" Then
            db.updateSoalFinal(txtIdNilai.Text, cbSiswaNilai.SelectedItem.Value, cbEkskulNilai.SelectedValue, txtMidNilai.Text, txtFinalNilai.Text)
            dataNilai()
            resetNilai()
            MsgBox("Update data success", MsgBoxStyle.Information)
            Exit Sub
        End If
    End Sub

    Private Sub DataGridViewNilai_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewNilai.CellMouseClick
        Dim i As Integer = DataGridViewNilai.CurrentRow.Index

        txtIdNilai.Text = DataGridViewNilai.Item(1, i).Value
        cbEkskulNilai.SelectedValue = DataGridViewNilai.Item(4, i).Value
        'cbSiswaNilai.SelectedItem.Value = DataGridViewNilai.Item(2, i).Value
        cbSiswaNilai.SelectedIndex = cbSiswaNilai.FindString(DataGridViewNilai.Item(3, i).Value).ToString()
        txtMidNilai.Text = DataGridViewNilai.Item(8, i).Value
        If Not DataGridViewNilai.Item(9, i).Value Is DBNull.Value Then
            txtFinalNilai.Text = DataGridViewNilai.Item(9, i).Value
        End If
    End Sub

    Private Sub DataGridViewNilai_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewNilai.CellEndEdit
        Dim i As Integer = DataGridViewNilai.CurrentRow.Index
        If e.ColumnIndex = 6 Then
            db.updateNilaiMid(DataGridViewNilai.Item(6, i).Value, DataGridViewNilai.Item(1, i).Value)
        ElseIf e.ColumnIndex = 7
            db.updateNilaiFinal(DataGridViewNilai.Item(7, i).Value, DataGridViewNilai.Item(1, i).Value)
        End If
    End Sub

    Private Sub DataGridViewNilai_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewNilai.CellClick
        Dim i As Integer = DataGridViewNilai.CurrentRow.Index
        If e.ColumnIndex = 12 Then
            Try
                Process.Start(DataGridViewNilai.Item(10, i).Value)
            Catch ex As Exception
                MsgBox("Jawaban belum diupload", MsgBoxStyle.Exclamation)
            End Try
        ElseIf e.ColumnIndex = 13
            Try
                Process.Start(DataGridViewNilai.Item(11, i).Value)
            Catch ex As Exception
                MsgBox("Jawaban belum diupload", MsgBoxStyle.Exclamation)
            End Try
        End If
    End Sub
#End Region
End Class
