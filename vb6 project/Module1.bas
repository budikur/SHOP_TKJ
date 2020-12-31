Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public RS_MASTER_BRG As New ADODB.Recordset
Public RS_KELUAR As New ADODB.Recordset
Public RS_STOK As New ADODB.Recordset
Public RS_BELI As New ADODB.Recordset
Public RS_JUAL As New ADODB.Recordset
Public RS_MODAL As New ADODB.Recordset
Public path_db As String
Public CARI As Boolean
Public KD_BRG As String

Public Sub CLEAR()
    MASTER_BRG.KD_BRG.Text = ""
    MASTER_BRG.NAMA_BRG.Text = ""
    MASTER_BRG.JNS_BRG.Text = ""
    MASTER_BRG.HRG_BELI.Text = ""
    MASTER_BRG.HRG_JUAL.Text = ""
    MASTER_BRG.JML_BAIK.Text = ""
    MASTER_BRG.JML_RUSAK.Text = ""
End Sub

Public Sub DEFAULT()
    MASTER_BRG.CMD_BARU.Enabled = True
    MASTER_BRG.CMD_UBAH.Enabled = True
    MASTER_BRG.CMD_SIMPAN.Enabled = False
    MASTER_BRG.CMD_HAPUS.Enabled = True
    MASTER_BRG.CMD_HAPUS.Caption = "HAPUS"
    MASTER_BRG.KD_BRG.Enabled = False
    MASTER_BRG.NAMA_BRG.Enabled = False
    MASTER_BRG.JNS_BRG.Enabled = False
    MASTER_BRG.HRG_BELI.Enabled = False
    MASTER_BRG.HRG_JUAL.Enabled = False
    MASTER_BRG.JML_BAIK.Enabled = False
    MASTER_BRG.JML_RUSAK.Enabled = False
    MASTER_BRG.DataGrid1.Columns(1).Width = 3000
    MASTER_BRG.DataGrid1.Columns(3).Width = 1000
    MASTER_BRG.DataGrid1.Columns(4).Width = 1000
    MASTER_BRG.DataGrid1.Columns(5).Width = 1000
    MASTER_BRG.DataGrid1.Columns(6).Width = 1000
    MASTER_BRG.DataGrid1.Columns(3).Alignment = dbgRight
    MASTER_BRG.DataGrid1.Columns(4).Alignment = dbgRight
    MASTER_BRG.DataGrid1.Columns(5).Alignment = dbgRight
    MASTER_BRG.DataGrid1.Columns(6).Alignment = dbgRight
    MASTER_BRG.JML_RECORD.Text = RS_MASTER_BRG.RecordCount
End Sub

Public Sub DISPLAY()
    If Module1.RS_MASTER_BRG.BOF Then
        Module1.RS_MASTER_BRG.MoveFirst
    End If
    If Module1.RS_MASTER_BRG.EOF Then
        Module1.RS_MASTER_BRG.MoveLast
    End If
    MASTER_BRG.KD_BRG.Text = RS_MASTER_BRG(0)
    MASTER_BRG.NAMA_BRG.Text = RS_MASTER_BRG(1)
'    JNS_BRG.Text = RS_MASTER_BRG(2)
    MASTER_BRG.HRG_BELI.Text = RS_MASTER_BRG(3)
    MASTER_BRG.HRG_JUAL.Text = RS_MASTER_BRG(4)
    MASTER_BRG.JML_BAIK.Text = RS_MASTER_BRG(5)
    MASTER_BRG.JML_RUSAK.Text = RS_MASTER_BRG(6)
End Sub

Public Sub NEW_RECORD()
    RS_MASTER_BRG.AddNew
    RS_MASTER_BRG(0) = MASTER_BRG.KD_BRG.Text
    RS_MASTER_BRG(1) = MASTER_BRG.NAMA_BRG.Text
    RS_MASTER_BRG(2) = MASTER_BRG.JNS_BRG.Text
    If MASTER_BRG.HRG_BELI.Text = "" Then
        RS_MASTER_BRG(3) = 0
    Else
        RS_MASTER_BRG(3) = MASTER_BRG.HRG_BELI.Text
    End If
    If MASTER_BRG.HRG_JUAL.Text = "" Then
        RS_MASTER_BRG(4) = 0
    Else
        RS_MASTER_BRG(4) = MASTER_BRG.HRG_JUAL.Text
    End If
    If MASTER_BRG.JML_BAIK.Text = "" Then
        RS_MASTER_BRG(5) = 0
    Else
        RS_MASTER_BRG(5) = MASTER_BRG.JML_BAIK.Text
    End If
    If MASTER_BRG.JML_RUSAK.Text = "" Then
        RS_MASTER_BRG(6) = 0
    Else
        RS_MASTER_BRG(6) = MASTER_BRG.JML_RUSAK.Text
    End If
    
    RS_MASTER_BRG(7) = "EXIST"
    RS_MASTER_BRG.UPDATE
    RS_MASTER_BRG.Requery
    Set MASTER_BRG.DataGrid1.DataSource = RS_MASTER_BRG
End Sub

Public Sub SEARCHING()
    RS_MASTER_BRG.MoveFirst
    RS_MASTER_BRG.Find "KD_BRG='" & MASTER_BRG.KD_BRG.Text & "'"
    If RS_MASTER_BRG.EOF Then
        CARI = False
    Else
        CARI = True
    End If
End Sub

Public Sub SORTING()
    RS_MASTER_BRG.Sort = "KD_BRG"
    RS_MASTER_BRG.MoveFirst
End Sub

Public Sub UPDATE()
    RS_MASTER_BRG(0) = MASTER_BRG.KD_BRG.Text
    RS_MASTER_BRG(1) = MASTER_BRG.NAMA_BRG.Text
    RS_MASTER_BRG(2) = MASTER_BRG.JNS_BRG.Text
    If MASTER_BRG.HRG_BELI.Text = "" Then
        RS_MASTER_BRG(3) = 0
    Else
        RS_MASTER_BRG(3) = MASTER_BRG.HRG_BELI.Text
    End If
    If MASTER_BRG.HRG_JUAL.Text = "" Then
        RS_MASTER_BRG(4) = 0
    Else
        RS_MASTER_BRG(4) = MASTER_BRG.HRG_JUAL.Text
    End If
    If MASTER_BRG.JML_BAIK.Text = "" Then
        RS_MASTER_BRG(5) = 0
    Else
        RS_MASTER_BRG(5) = MASTER_BRG.JML_BAIK.Text
    End If
    If MASTER_BRG.JML_RUSAK.Text = "" Then
        RS_MASTER_BRG(6) = 0
    Else
        RS_MASTER_BRG(6) = MASTER_BRG.JML_RUSAK.Text
    End If
    RS_MASTER_BRG(7) = "EXIST"
    RS_MASTER_BRG.UPDATE
    RS_MASTER_BRG.Requery
    Set MASTER_BRG.DataGrid1.DataSource = RS_MASTER_BRG
End Sub

