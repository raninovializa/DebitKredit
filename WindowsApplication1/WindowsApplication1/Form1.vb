Imports System.Data.OleDb
Public Class Form1
    Sub kosong()
        TextBox1.Clear()
        TextBox4.Clear()
        TextBox3.Clear()
        TextBox2.Clear()
    End Sub
    Sub isi()
        TextBox4.Clear()
        TextBox4.Focus()
        TextBox3.Clear()
        TextBox3.Focus()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilJenis()
        da = New OleDbDataAdapter("select * from jenis", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "jenis")
        DataGridView1.DataSource = ds.Tables("jenis")
        DataGridView1.Refresh()
    End Sub
    Sub AturGrid()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 200
        DataGridView1.Columns(0).HeaderText = "Kode Akun"
        DataGridView1.Columns(0).HeaderText = "Kode Akun"
    End Sub
    Private Sub DataJenis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilJenis()
        Call kosong()
        Call AturGrid()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("data belum lengkap....!!")
            TextBox1.Focus()
            Exit Sub
        Else
            cmd = New OleDbCommand("select * from jenis where Kode_Akun='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim simpan As String = "insert into jenis (Kode_Akun, Nama_Akun, Debet, Kredit)values" & _
                "('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                cmd = New OleDbCommand(simpan, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("simpan data sukses....!", MsgBoxStyle.Information, "perhatian")
            End If
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("kode akun belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim ubah As String = "update jenis set Nama_Akun = '" & TextBox2.Text & "' , Debet = '" & TextBox3.Text & "' , kredit = '" & TextBox4.Text & "' Where Kode_Akun = '" & TextBox1.Text & "'"
            cmd = New OleDbCommand(ubah, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("ubah data sukses...!!", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call kosong()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim hapus As String = "delete from jenis Where Kode_Akun = '" & TextBox1.Text & "' or Nama_Akun = '" & TextBox2.Text & "'"
        cmd = New OleDbCommand(hapus, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("hapus data sukses...!!", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value
        TextBox2.Text = DataGridView1.CurrentRow.Cells(1).Value
        TextBox3.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString()
        TextBox4.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString()
    End Sub

End Class