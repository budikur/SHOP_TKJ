VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form STOK_VIEW_2 
   Caption         =   "STOK BARANG"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "MASTER_VIEW.frx":0000
      Left            =   1440
      List            =   "MASTER_VIEW.frx":000D
      TabIndex        =   3
      Text            =   "NAMA_BRG"
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "MASTER_VIEW.frx":002E
      Left            =   4680
      List            =   "MASTER_VIEW.frx":003B
      TabIndex        =   2
      Text            =   "NAMA_BRG"
      Top             =   120
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGRID 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6588
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "SEARCH"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "SORT BY"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "BY"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "STOK_VIEW_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ASD As String

Private Sub Combo2_Click()
    RS_STOK.Sort = Combo2.Text
End Sub

Private Sub FGRID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If PEMBELIAN.Visible = True Then
            PEMBELIAN.KONEKSI
            PEMBELIAN.FGRID.TextMatrix(PEMBELIAN.asd_row, 1) = FGRID.TextMatrix(FGRID.Row, 1)
            PEMBELIAN.FGRID.TextMatrix(PEMBELIAN.asd_row, 2) = FGRID.TextMatrix(FGRID.Row, 2)
'            PEMBELIAN.FGRID.TextMatrix(PEMBELIAN.asd_row, 1) = FGRID.TextMatrix(FGRID.Row, 3)
'            PEMBELIAN.FGRID.TextMatrix(PEMBELIAN.asd_row, 2) = FGRID.TextMatrix(FGRID.Row, 4)
            PEMBELIAN.SetFocus
            PEMBELIAN.FGRID.Col = PEMBELIAN.asd_col + 2
            PEMBELIAN.FGRID.Row = PEMBELIAN.asd_row
            PEMBELIAN.EDIT_FGRID
        ElseIf PENJUALAN.Visible = True Then
            PENJUALAN.J_KONEKSI
'            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 1) = FGRID.TextMatrix(FGRID.Row, 1)
'            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 2) = FGRID.TextMatrix(FGRID.Row, 2)
'            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 4) = FGRID.TextMatrix(FGRID.Row, 5)
            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 1) = FGRID.TextMatrix(FGRID.Row, 3)
            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 2) = FGRID.TextMatrix(FGRID.Row, 4)
            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 4) = FGRID.TextMatrix(FGRID.Row, 7)
            PENJUALAN.FGRID.TextMatrix(PENJUALAN.dsa_row, 5) = FGRID.TextMatrix(FGRID.Row, 6)
            PENJUALAN.SetFocus
            PENJUALAN.FGRID.Col = PENJUALAN.dsa_col + 2
            PENJUALAN.FGRID.Row = PENJUALAN.dsa_row
            PENJUALAN.J_EDIT_FGRID
        ElseIf PENGELUARAN.Visible = True Then
            PENGELUARAN.TXT_JENIS.Visible = True
            PENGELUARAN.TXT_NILAI.Visible = True
            PENGELUARAN.TXT_JENIS = "PENGAMBILAN : " & FGRID.TextMatrix(FGRID.Row, 2)
            PENGELUARAN.TXT_NILAI = FGRID.TextMatrix(FGRID.Row, 4)
            kd_brg = FGRID.TextMatrix(FGRID.Row, 1)
            PENGELUARAN.KONEKSI
            PENGELUARAN.TXT_NILAI.SetFocus
        End If
        Unload Me
    ElseIf KeyAscii = 27 Then
        If PEMBELIAN.Visible = True Then
            PEMBELIAN.KONEKSI
            PEMBELIAN.SetFocus
        ElseIf PENJUALAN.Visible = True Then
            PENJUALAN.J_KONEKSI
            PENJUALAN.SetFocus
        ElseIf PENGELUARAN.Visible = True Then
            PENGELUARAN.KONEKSI
        End If
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then con.Close
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    
    RS_STOK.Open "select KD_BRG,NAMA_BRG,JENIS_BRG,HRG_BELI,HRG_JUAL,JML_BAIK,JML_RUSAK from MASTER_BRG_JASA WHERE HAPUS='EXIST'", con, adOpenKeyset, adLockOptimistic
'    RS_STOK.Open "select ID_BELI,TGL_BELI,KD_BRG,NAMA_BRG,JML_BELI,HRG_BELI,HRG_JUAL from PEMBELIAN", con, adOpenKeyset, adLockOptimistic
    Set FGRID.DataSource = RS_STOK
    RS_STOK.Sort = "NAMA_BRG"
'    RS_STOK.Sort = "TGL_BELI"
    
    FGRID.ColWidth(2) = 3000
    RS_STOK.MoveLast
    
'    FGRID.Columns(1).Width = 3000
'    DataGrid1.Columns(3).Width = 1000
'    DataGrid1.Columns(4).Width = 1000
'    DataGrid1.Columns(3).Alignment = dbgRight
'    DataGrid1.Columns(4).Alignment = dbgRight
End Sub

Private Sub Text1_Change()
'    RS_STOK.Filter = Combo1.Text & "='" & Text1.Text & "'"
'    rsFilter.Filter = "Kode = 'AAA'"
'    Adodc1.RecordSource = "select * from data where nama like'%" & Text1.Text & "%'"
    RS_STOK.Close
    RS_STOK.Open "select * from MASTER_BRG_JASA where " & Combo1.Text & " like'%" & Text1.Text & "%'", con, adOpenKeyset, adLockOptimistic
'    RS_STOK.Open "select * from PEMBELIAN where " & Combo1.Text & " like'%" & Text1.Text & "%'", con, adOpenKeyset, adLockOptimistic
    RS_STOK.Requery
    Set FGRID.DataSource = RS_STOK
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FGRID.SetFocus
    ElseIf KeyAscii = 27 Then
        If PEMBELIAN.Visible = True Then
            PEMBELIAN.KONEKSI
            PEMBELIAN.SetFocus
            Unload Me
        ElseIf PENJUALAN.Visible = True Then
            PENJUALAN.J_KONEKSI
            PENJUALAN.SetFocus
            Unload Me
        End If
    End If
End Sub
