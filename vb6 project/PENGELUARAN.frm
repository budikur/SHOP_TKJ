VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PENGELUARAN 
   Caption         =   "PENGELUARAN"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form3"
   ScaleHeight     =   7455
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39948
   End
   Begin VB.TextBox TXT_RUSAK 
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "PENGELUARAN.frx":0000
      Left            =   2160
      List            =   "PENGELUARAN.frx":000A
      TabIndex        =   1
      Text            =   "PILIH SALAH SATU"
      Top             =   1320
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9551
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
   Begin VB.TextBox TXT_NILAI 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox TXT_JENIS 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox TXT_ID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "PERALATAN RUSAK"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "KETERANGAN"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "NILAI PENGELUARAN"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "JENIS PENGELUARAN"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "TGL PENGELUARAN"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ID PENGELUARAN"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "PENGELUARAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DISPLAY()
    DataGrid1.Columns(0).Width = 700
    DataGrid1.Columns(2).Width = 5000
    DataGrid1.Columns(3).Width = 2000
    DataGrid1.Columns(3).Alignment = dbgRight
End Sub

Private Sub Combo1_GotFocus()
    TXT_RUSAK.Visible = False
    TXT_JENIS.Visible = False
    TXT_NILAI.Visible = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Combo1.Text = "NON BARANG" Then
            TXT_RUSAK.Visible = False
            TXT_JENIS.Visible = True
            TXT_NILAI.Visible = True
            TXT_JENIS.SetFocus
        ElseIf Combo1.Text = "BARANG" Then
            TXT_RUSAK.Visible = True
            TXT_RUSAK.SetFocus
            TXT_JENIS.Visible = False
            TXT_NILAI.Visible = False
        Else
            TXT_RUSAK.Visible = False
            TXT_JENIS.Visible = False
            TXT_NILAI.Visible = False
            Combo1.SetFocus
        End If
    End If
End Sub

'Private Sub Form_Load()
'    KONEKSI
'    TXT_ID.Text = "K-" & RS_KELUAR.RecordCount + 1
'    RS_KELUAR.Sort = "TGL_PENGELUARAN"
'    RS_KELUAR.MoveLast
'    DISPLAY
'End Sub
Private Sub Form_Load()
    KONEKSI
    RS_KELUAR.MoveFirst
    TEMP = "K-" & RS_KELUAR.RecordCount + 1
    RS_KELUAR.Find "ID_PENGELUARAN='" & TEMP & "'"
    i = 1
    While RS_KELUAR.EOF = False
        i = i + 1
        RS_KELUAR.MoveFirst
        TEMP = "K-" & RS_KELUAR.RecordCount + i
        RS_KELUAR.Find "ID_PENGELUARAN='" & TEMP & "'"
    Wend
    
    TXT_ID.Text = "K-" & RS_KELUAR.RecordCount + i
    RS_KELUAR.Sort = "TGL_PENGELUARAN"
    RS_KELUAR.MoveLast
    DISPLAY
End Sub

Public Sub KONEKSI()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then con.Close
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    
    RS_KELUAR.Open "select * from pengeluaran", con, adOpenKeyset, adLockOptimistic
    RS_MASTER_BRG.Open "select * from MASTER_BRG_JASA", con, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RS_KELUAR
End Sub

Private Sub TXT_JENIS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TXT_NILAI.SetFocus
    End If
End Sub

Private Sub TXT_NILAI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RS_KELUAR.AddNew
        RS_KELUAR(0) = TXT_ID.Text
        RS_KELUAR(1) = DTPicker1.Value
        
        If TXT_RUSAK.Visible = True Then
            RS_KELUAR(2) = TXT_RUSAK.Text
            RS_KELUAR(4) = TXT_JENIS.Text
            '=====================================
            'UPDATE MASTER BRG
            '=====================================
            RS_MASTER_BRG.Sort = "KD_BRG"
            RS_MASTER_BRG.MoveFirst
            RS_MASTER_BRG.Find "KD_BRG='" & KD_BRG & "'"
            RS_MASTER_BRG(5) = RS_MASTER_BRG(5) - 1
            RS_MASTER_BRG.UPDATE
            RS_MASTER_BRG.Requery
        Else
            RS_KELUAR(2) = TXT_JENIS.Text
        End If
      
        RS_KELUAR(3) = TXT_NILAI.Text
        RS_KELUAR.UPDATE
        RS_KELUAR.Requery
        RS_KELUAR.Sort = "TGL_PENGELUARAN"
        RS_KELUAR.MoveLast
        Set DataGrid1.DataSource = RS_KELUAR
        
        TXT_ID.Text = "K-" & RS_KELUAR.RecordCount + 1
        TXT_RUSAK.Text = ""
        TXT_JENIS.Text = ""
        TXT_NILAI.Text = ""
        DTPicker1.SetFocus
        DISPLAY
    End If
End Sub

Private Sub TXT_RUSAK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        STOK_VIEW_2.Show
    End If
End Sub

