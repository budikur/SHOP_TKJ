VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form LABA_RUGI 
   Caption         =   "LABA RUGI DAN PERUBAHAN MODAL"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD_FRESH 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   11880
      TabIndex        =   27
      Top             =   120
      Width           =   3015
   End
   Begin MSACAL.Calendar CLD_AKR 
      Height          =   2295
      Left            =   5280
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   3
      Day             =   13
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXT_AKR 
      Height          =   285
      Left            =   5280
      TabIndex        =   25
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TXT_AWL 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   330
      Width           =   1575
   End
   Begin MSACAL.Calendar CLD_AWL 
      Height          =   2295
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   3
      Day             =   13
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXT_M_AWL 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11400
      TabIndex        =   19
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox TXT_LR2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11400
      TabIndex        =   18
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox TXT_M_AKR 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11400
      TabIndex        =   17
      Top             =   7200
      Width           =   2415
   End
   Begin VB.TextBox TXT_LR1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Rp""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox TXT_KLR 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox TXT_JSA 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox TXT_BRG 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   6600
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DG_KELUAR 
      Height          =   2295
      Left            =   7800
      TabIndex        =   4
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
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
   Begin MSDataGridLib.DataGrid DG_JUAL 
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
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
   Begin VB.CommandButton CMD_PROSES 
      Caption         =   "PROSES"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DG_MODAL 
      Height          =   2415
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   4260
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
   Begin VB.Label Label12 
      Caption         =   "SAMPAI TANGGAL :"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   390
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "MULAI TANGGAL :"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "MODAL AWAL BULAN      ="
      Height          =   255
      Left            =   9120
      TabIndex        =   22
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "BESARNYA LABA/RUBI   ="
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "MODAL AKHIR BULAN     ="
      Height          =   255
      Left            =   9120
      TabIndex        =   20
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "BESARNYA LABA/RUGI        ="
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "TOTAL PENGELUARAN        ="
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "LABA PELAYANAN JASA       ="
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "LABA PENJUALAN BARANG ="
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "PERUBAHAN MODAL :"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "PENGELUARAN :"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "PENJUALAN BARANG DAN JASA :"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "LABA_RUGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TEMP As Currency
Dim TEMP2 As Currency
Dim DTEMP As Date

Private Sub CLD_AKR_Click()
    TXT_AKR.Text = CLD_AKR.Value
    
    CLD_AKR.Visible = False
    CMD_PROSES.SetFocus
End Sub

Private Sub CLD_AKR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD_AKR_Click
    ElseIf KeyAscii = 27 Then
        CLD_AKR.Visible = False
        TXT_AKR.SetFocus
    End If
End Sub

Private Sub CLD_AWL_Click()
    TXT_AWL.Text = CLD_AWL.Value

    CLD_AWL.Visible = False
    TXT_AKR.SetFocus
End Sub

Private Sub CLD_AWL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD_AWL_Click
    ElseIf KeyAscii = 27 Then
        CLD_AWL.Visible = False
        TXT_AWL.SetFocus
    End If
End Sub

Private Sub CMD_FRESH_Click()
    Form_Load
End Sub

Private Sub CMD_PROSES_Click()
    On Error Resume Next
    '=================================
    'FILTER DATA
    '=================================
    RS_JUAL.Filter = "tgl_jual >='" & TXT_AWL.Text & "' And tgl_jual <='" & TXT_AKR.Text & "'"
    RS_KELUAR.Filter = "tgl_pengeluaran >='" & TXT_AWL.Text & "' And tgl_pengeluaran <='" & TXT_AKR.Text & "'"
    RS_MODAL.Filter = "tgl_harian >='" & TXT_AWL.Text & "' And tgl_harian <='" & TXT_AKR.Text & "'"
    '=================================
    'PENGELUARAN
    '=================================
    TEMP = 0
    RS_KELUAR.MoveFirst
    While RS_KELUAR.EOF = False
        TEMP = TEMP + RS_KELUAR(3)
        RS_KELUAR.MoveNext
    Wend
    TXT_KLR.Text = TEMP
    '=================================
    'LABA BRG & JASA
    '=================================
    TEMP = 0
    TEMP2 = 0
    RS_JUAL.MoveFirst
    While RS_JUAL.EOF = False
        If Mid(RS_JUAL(2), 1, 3) = "BRG" Then
            TEMP = TEMP + RS_JUAL(7)
        ElseIf Mid(RS_JUAL(2), 1, 3) = "JSA" Then
            TEMP2 = TEMP2 + RS_JUAL(7)
        End If
        RS_JUAL.MoveNext
    Wend
    TXT_BRG.Text = TEMP
    TXT_JSA.Text = TEMP2
    '=================================
    'LABA TOTAL
    '=================================
    TEMP = 0
    RS_MODAL.MoveFirst
    TEMP2 = RS_MODAL(3)
    While RS_MODAL.EOF = False
        TEMP = TEMP + RS_MODAL(2)
        RS_MODAL.MoveNext
    Wend
    TXT_LR1.Text = TEMP
    TXT_LR2.Text = TEMP
    TXT_M_AWL = TEMP2
    TXT_M_AKR = TEMP2 + TEMP
End Sub

Private Sub CMD_PROSES_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        TXT_AKR.Text = ""
        TXT_AKR.SetFocus
    End If
End Sub

Private Sub Form_Load()
    KONEKSI
    INISIALISASI_LABA
    UPDATE_LABA
    UPDATE_MODAL
End Sub

Private Sub INISIALISASI_LABA()
    RS_MODAL.MoveFirst
    While RS_MODAL.EOF = False
        RS_MODAL(2) = 0
        RS_MODAL.MoveNext
    Wend
End Sub

Private Sub KONEKSI()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then
        con.Close
    End If
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    RS_JUAL.Open "select * from penjualan", con, adOpenKeyset, adLockOptimistic
    RS_KELUAR.Open "select * from pengeluaran", con, adOpenKeyset, adLockOptimistic
    RS_MODAL.Open "select * from KEUANGAN", con, adOpenKeyset, adLockOptimistic
    Set DG_JUAL.DataSource = RS_JUAL
    Set DG_KELUAR.DataSource = RS_KELUAR
    Set DG_MODAL.DataSource = RS_MODAL
    RS_JUAL.Sort = "TGL_JUAL"
    RS_KELUAR.Sort = "TGL_PENGELUARAN"
    RS_MODAL.Sort = "TGL_HARIAN"
End Sub

Private Sub UPDATE_LABA()
    On Error Resume Next
    '================================
    'DARI TABEL PENJUALAN
    '================================
    RS_JUAL.MoveFirst
    RS_MODAL.MoveFirst
    TEMP = 0
    DTEMP = RS_JUAL(1)
    While RS_JUAL.EOF = False
        If RS_JUAL(1) = DTEMP Then
            TEMP = TEMP + RS_JUAL(7)
        Else
            RS_MODAL.Find "TGL_HARIAN='" & DTEMP & "'"
            If RS_MODAL.EOF = True Then
                RS_MODAL.AddNew
                RS_MODAL(0) = "LR-" & RS_MODAL.RecordCount + 1
                RS_MODAL(1) = DTEMP
                RS_MODAL(2) = TEMP
            Else
                RS_MODAL(2) = TEMP
            End If
            RS_MODAL.UPDATE
            RS_MODAL.Requery
            TEMP = RS_JUAL(7)
        End If
        DTEMP = RS_JUAL(1)
        RS_JUAL.MoveNext
    Wend
    RS_MODAL.MoveFirst
    RS_MODAL.Find "TGL_HARIAN='" & DTEMP & "'"
    If RS_MODAL.EOF = True Then
        RS_MODAL.AddNew
        RS_MODAL(0) = "LR-" & RS_MODAL.RecordCount + 1
        RS_MODAL(1) = DTEMP
        RS_MODAL(2) = TEMP
    Else
        RS_MODAL(2) = TEMP
    End If
    RS_MODAL.UPDATE
    RS_MODAL.Requery
    '================================
    'DARI TABEL PENGELUARAN
    '================================
    RS_KELUAR.MoveFirst
    RS_MODAL.MoveFirst
    TEMP = 0
    DTEMP = RS_KELUAR(1)
    While RS_KELUAR.EOF = False
        If RS_KELUAR(1) = DTEMP Then
            TEMP = TEMP + RS_KELUAR(3)
        Else
            RS_MODAL.Find "TGL_HARIAN='" & DTEMP & "'"
            If RS_MODAL.EOF = True Then
                RS_MODAL.AddNew
                RS_MODAL(0) = "LR-" & RS_MODAL.RecordCount + 1
                RS_MODAL(1) = DTEMP
                RS_MODAL(2) = -TEMP
            Else
                RS_MODAL(2) = RS_MODAL(2) - TEMP
            End If
            RS_MODAL.UPDATE
            RS_MODAL.Requery
            TEMP = RS_KELUAR(3)
        End If
        DTEMP = RS_KELUAR(1)
        RS_KELUAR.MoveNext
    Wend
    RS_MODAL.MoveFirst
    RS_MODAL.Find "TGL_HARIAN='" & DTEMP & "'"
    If RS_MODAL.EOF = True Then
        RS_MODAL.AddNew
        RS_MODAL(0) = "LR-" & RS_MODAL.RecordCount + 1
        RS_MODAL(1) = DTEMP
        RS_MODAL(2) = -TEMP
    Else
        RS_MODAL(2) = RS_MODAL(2) - TEMP
    End If
    RS_MODAL.UPDATE
    RS_MODAL.Requery
End Sub

Private Sub UPDATE_MODAL()
    On Error Resume Next
    RS_MODAL.MoveFirst
    While RS_MODAL.EOF = False
        RS_MODAL(4) = RS_MODAL(2) + RS_MODAL(3)
        TEMP = RS_MODAL(4)
        RS_MODAL.MoveNext
        If RS_MODAL.EOF = False Then
            RS_MODAL(3) = TEMP
        End If
    Wend
    RS_MODAL.UPDATE
    RS_MODAL.Requery
    RS_MODAL.MoveLast
    RS_JUAL.MoveLast
    RS_KELUAR.MoveLast
End Sub

Private Sub TXT_AKR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD_AKR.Visible = True
        CLD_AKR.SetFocus
    ElseIf KeyAscii = 27 Then
        TXT_AWL.Text = ""
        TXT_AWL.SetFocus
    End If
End Sub

Private Sub TXT_AWL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD_AWL.Visible = True
        CLD_AWL.SetFocus
    End If
End Sub
