Imports System.Data.SqlClient
Public Class personelAnasafya
    Dim baglanti As SqlConnection
    Dim sqlQuery As String
    Dim command As New SqlCommand
    Dim adapter As New SqlDataAdapter
    Dim table As New DataTable

    Dim gorevDetaylari(10) As String

    Private Sub personelAnasafya_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        baglanti = New SqlConnection("Data Source=10.40.125.46;User ID=Deneme;Password=123456;")
        baglanti.Open()


        PersonelBilgileriGetir()
        GorevBilgileriGetir()


    End Sub

    Private Sub btn_gorevTamamla_Click(sender As Object, e As EventArgs) Handles btn_gorevTamamla.Click
        Dim result1 As DialogResult
        If tb_gorevNotu.Text Then

        Else
            result1 = MessageBox.Show("Devam etmek ister misiniz?",
                                                      "Important Question",
                                                      MessageBoxButtons.YesNo)
        End If
        If result1 = DialogResult.Yes Then
            vt_yaz()
        End If

        GorevBilgileriGetir()
    End Sub



    Public Sub GorevBilgileriGetir()
        sqlQuery = "Select GorevAdı from Gorevler"
        command.Connection = baglanti
        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)

        For i = 0 To table.Rows.Count - 1
            lb_gorevler.Items.Add(table.Rows(i).Item(0).ToString())
        Next

        sqlQuery = "Select GorevAciklamasi from Gorevler"
        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)
        If table.Rows.Count > 10 Then
            ReDim Preserve gorevDetaylari(table.Rows.Count)
        End If
        For i = 0 To table.Rows.Count - 1
            gorevDetaylari(i) = table.Rows(i).Item(0).ToString()
        Next

    End Sub

    Private Sub lb_gorevler_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_gorevler.SelectedIndexChanged
        lbl_gorevDetay.Text = gorevDetaylari(lb_gorevler.SelectedIndex)
    End Sub

    Public Sub PersonelBilgileriGetir()
        sqlQuery = "Select * from Personeller where KullaniciAdi='dogukan58'"
        command.Connection = baglanti
        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)
        lbl_ad.Text = "Ad: " & table.Rows(0).Item(2).ToString
        lbl_soyad.Text = "Soyad: " & table.Rows(0).Item(3).ToString
        lbl_birim.Text = "Birim: " & table.Rows(0).Item(8).ToString
        lbl_tel.Text = "Telefon: " & table.Rows(0).Item(5).ToString
        sqlQuery = "Select * from Personeller where KullaniciID='4'"
        command.Connection = baglanti
        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)
        lbl_yoneticiTel.Text = "Tel: " & table.Rows(0).Item(5).ToString
        sqlQuery = "Select * from Gorevler where AtananPersonel='7'"
        command.Connection = baglanti
        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)
        lb_gorevler.Items.Add(table.Rows(0).Item(1).ToString)

    End Sub

    Public Sub vt_yaz()
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = "Data Source=10.40.125.46;User ID=Deneme;Password=123456;"
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "INSERT INTO Gorevler([field1], [field2]) VALUES([Value1], [Value2])"
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("kayit yerlestirmede hata olustu..." & ex.Message, "kayitlari ye")
        Finally
            con.Close()
        End Try
    End Sub

End Class