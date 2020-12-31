VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MASTER_BRG 
   Caption         =   "MASTER BARANG"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   23
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox JML_RECORD 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      DragMode        =   1  'Automatic
      Height          =   4935
      Left            =   360
      TabIndex        =   20
      Top             =   3120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMD_PREV 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton CMD_NEXT 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox JNS_BRG 
      Height          =   315
      ItemData        =   "MASTER_BRG.frx":0000
      Left            =   6720
      List            =   "MASTER_BRG.frx":0002
      TabIndex        =   17
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton CMD_SIMPAN 
      Caption         =   "SIMPAN"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton CMD_HAPUS 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton CMD_UBAH 
      Caption         =   "UBAH"
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CMD_BARU 
      Caption         =   "BARU"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox JML_RUSAK 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox JML_BAIK 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox HRG_JUAL 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox HRG_BELI 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox NAMA_BRG 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox KD_BRG 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "CARI BARANG :"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "JUMLAH RECORD"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   15240
      X2              =   0
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label7 
      Caption         =   "JUMLAH RUSAK"
      Height          =   255
      Left            =   10920
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "JUMLAH BAIK"
      Height          =   255
      Left            =   10920
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "HARGA JUAL"
      Height          =   255
      Left            =   10920
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "HARGA BELI"
      Height          =   255
      Left            =   10920
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "JENIS BARANG"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA BARANG"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "KODE BARANG"
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "MASTER_BRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMD_BARU_Click()
    CMD_BARU.Enabled = False
    CMD_UBAH.Enabled = False
    CMD_SIMPAN.Enabled = True
    CMD_HAPUS.Caption = "CANCEL"
    Module1.CLEAR
    
    NAMA_BRG.Enabled = True
    JNS_BRG.Enabled = True
    HRG_BELI.Enabled = True
    HRG_JUAL.Enabled = True
    JML_BAIK.Enabled = True
    JML_RUSAK.Enabled = True
    NAMA_BRG.SetFocus
    
    KD_BRG.Text = "BRG-" & RS_MASTER_BRG.RecordCount + 1
End Sub

Private Sub CMD_HAPUS_Click()
    If CMD_HAPUS.Caption = "CANCEL" Then
        Module1.DEFAULT
    Else
        Module1.RS_MASTER_BRG(7) = "HAPUS"
        Module1.RS_MASTER_BRG.UPDATE
        Module1.RS_MASTER_BRG.Requery
        Set MASTER_BRG.DataGrid1.DataSource = RS_MASTER_BRG
    End If
    Module1.DISPLAY
End Sub

Private Sub CMD_NEXT_Click()
    Module1.RS_MASTER_BRG.MoveNext
    Module1.DISPLAY
End Sub

Private Sub CMD_PREV_Click()
    Module1.RS_MASTER_BRG.MovePrevious
    Module1.DISPLAY
End Sub

Private Sub CMD_SIMPAN_Click()
    Module1.SEARCHING
    If CARI = True Then
        YES_NO.Show
    Else
        Module1.NEW_RECORD
        Module1.SORTING
        Module1.DISPLAY
        Module1.DEFAULT
    End If
'    CMD_BARU.SetFocus
End Sub

Private Sub CMD_UBAH_Click()
    CMD_BARU.Enabled = False
    CMD_UBAH.Enabled = False
    CMD_SIMPAN.Enabled = True
    CMD_HAPUS.Caption = "CANCEL"
    
    KD_BRG.Enabled = True
    NAMA_BRG.Enabled = True
    JNS_BRG.Enabled = True
    HRG_BELI.Enabled = True
    HRG_JUAL.Enabled = True
    JML_BAIK.Enabled = True
    JML_RUSAK.Enabled = True
    KD_BRG.SetFocus
End Sub

Private Sub DataGrid1_Click()
    Module1.DISPLAY
End Sub

Private Sub Form_Load()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then con.Close
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    '=============================================
    'MENYIMPAN SELURUH RECORD DI VARIABEL RS_MASTER_BRG
    '=============================================
    RS_MASTER_BRG.Open "select * from master_brg_jasa", con, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RS_MASTER_BRG
    
    Module1.DEFAULT
    Module1.SORTING
    Module1.DISPLAY
End Sub

Private Sub HRG_BELI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HRG_JUAL.SetFocus
    End If
End Sub

Private Sub HRG_JUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        JML_BAIK.SetFocus
    End If
End Sub

Private Sub JML_BAIK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        JML_RUSAK.SetFocus
    End If
End Sub

Private Sub JML_RUSAK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CMD_SIMPAN.SetFocus
    End If
End Sub

Private Sub JNS_BRG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HRG_BELI.SetFocus
    End If
End Sub

Private Sub KD_BRG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NAMA_BRG.SetFocus
    End If
    CMD_SIMPAN.Enabled = True
End Sub

Private Sub NAMA_BRG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        JNS_BRG.SetFocus
    End If
End Sub

Private Sub Text1_Change()
'    RS_STOK.Filter = Combo1.Text & "='" & Text1.Text & "'"
'    rsFilter.Filter = "Kode = 'AAA'"
'    Adodc1.RecordSource = "select * from data where nama like'%" & Text1.Text & "%'"
    RS_MASTER_BRG.Close
    RS_MASTER_BRG.Open "select * from MASTER_BRG_JASA where NAMA_BRG like'%" & Text1.Text & "%'", con, adOpenKeyset, adLockOptimistic
    RS_MASTER_BRG.Requery
    Set DataGrid1.DataSource = RS_MASTER_BRG
End Sub

