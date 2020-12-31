VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form STOK_BRG 
   Caption         =   "STOK BARANG"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "STOK_BRG.frx":0000
      Left            =   4920
      List            =   "STOK_BRG.frx":000D
      TabIndex        =   5
      Text            =   "NAMA_BRG"
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "STOK_BRG.frx":002E
      Left            =   1680
      List            =   "STOK_BRG.frx":003B
      TabIndex        =   4
      Text            =   "NAMA_BRG"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6588
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "BY"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "SORT BY"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "SEARCH"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "STOK_BRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ASD As String

Private Sub Combo2_Click()
    RS_STOK.Sort = Combo2.Text
End Sub

Private Sub Form_Load()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then con.Close
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    
    RS_STOK.Open "select KD_BRG,NAMA_BRG,JENIS_BRG,JML_BAIK,JML_RUSAK from MASTER_BRG_JASA WHERE HAPUS='EXIST'", con, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RS_STOK
    RS_STOK.Sort = "NAMA_BRG"
    DataGrid1.Columns(1).Width = 3000
    DataGrid1.Columns(3).Width = 1000
    DataGrid1.Columns(4).Width = 1000
    DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(4).Alignment = dbgRight
End Sub

Private Sub Text1_Change()
'    RS_STOK.Filter = Combo1.Text & "='" & Text1.Text & "'"
'    rsFilter.Filter = "Kode = 'AAA'"
'    Adodc1.RecordSource = "select * from data where nama like'%" & Text1.Text & "%'"
    RS_STOK.Close
    RS_STOK.Open "select KD_BRG,NAMA_BRG,JENIS_BRG,JML_BAIK,JML_RUSAK from MASTER_BRG_JASA where " & Combo1.Text & " like'%" & Text1.Text & "%'", con, adOpenKeyset, adLockOptimistic
    RS_STOK.Requery
    Set DataGrid1.DataSource = RS_STOK
End Sub
