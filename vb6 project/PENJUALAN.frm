VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PENJUALAN 
   Caption         =   "PENJUALAN"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form5"
   ScaleHeight     =   7800
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CELL 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TXT_TGL 
      Height          =   405
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CMD_INPUT 
      Caption         =   "INPUT DATA"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12360
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TXT_TOTAL 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   12600
      TabIndex        =   2
      Text            =   "0"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton CMD_NEW 
      Caption         =   "NEW DATA"
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   600
      TabIndex        =   4
      Top             =   4080
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   5953
      _Version        =   393216
      Enabled         =   -1  'True
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
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1335,118
         EndProperty
      EndProperty
   End
   Begin MSACAL.Calendar CLD 
      Height          =   2295
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      _Version        =   524288
      _ExtentX        =   7011
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2014
      Month           =   5
      Day             =   1
      DayLength       =   1
      MonthLength     =   2
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
   Begin MSFlexGridLib.MSFlexGrid FGRID 
      Height          =   2655
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      Cols            =   7
   End
   Begin VB.Label Label1 
      Caption         =   "TANGGAL"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "TABEL PENJUALAN"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "TOTAL PENJUALAN ="
      Height          =   255
      Left            =   10800
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "PENJUALAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dsa_col As String
Public dsa_row As String
Dim TEMP As Currency

Private Sub CLD_Click()
    TXT_TGL.Text = CLD.Value
    CLD.Visible = False
    FGRID.Rows = FGRID.Rows + 1
    FGRID.SetFocus
End Sub

Private Sub CLD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD_Click
    End If
End Sub

Private Sub CMD_INPUT_Click()
    On Error Resume Next
    If FGRID.TextMatrix(1, 1) <> "" Then
'        TXT_ID.Text = "K-" & RS_KELUAR.RecordCount + 1
        For i = 1 To (FGRID.Rows - 2)
            RS_JUAL.AddNew
            RS_JUAL(0) = "PJ-" & RS_JUAL.RecordCount + 1
            RS_JUAL(1) = TXT_TGL.Text
            RS_JUAL(2) = FGRID.TextMatrix(i, 1)
            RS_JUAL(3) = FGRID.TextMatrix(i, 2)
            RS_JUAL(4) = FGRID.TextMatrix(i, 3)
            RS_JUAL(5) = FGRID.TextMatrix(i, 4)
            RS_JUAL(6) = FGRID.TextMatrix(i, 5)
'            =====================================
'            UPDATE MASTER BRG
'            =====================================
            RS_MASTER_BRG.Sort = "KD_BRG"
            RS_MASTER_BRG.MoveFirst
            RS_MASTER_BRG.Find "KD_BRG='" & FGRID.TextMatrix(i, 1) & "'"
'            HRG_BELI=RS_MASTER_BRG(3)
'            JML_BAIK=RS_MASTER_BRG(5)
'            JML_JUAL=FGRID.TextMatrix(i, 3)
'            HRG_JUAL=FGRID.TextMatrix(i, 4)
            If Left(RS_MASTER_BRG(0), 3) = "BRG" Then
                RS_MASTER_BRG(5) = RS_MASTER_BRG(5) - FGRID.TextMatrix(i, 3)
            End If
'            RS_JUAL(6) = RS_MASTER_BRG(3)
            RS_JUAL(7) = (RS_JUAL(5) - RS_JUAL(6)) * RS_JUAL(4)
            RS_JUAL.UPDATE
            RS_JUAL.Requery
            Set DataGrid1.DataSource = RS_JUAL
            RS_MASTER_BRG.UPDATE
            RS_MASTER_BRG.Requery
        Next
        CMD_INPUT.Enabled = False
        RS_JUAL.MoveLast
        CMD_NEW.SetFocus
    End If
End Sub

Private Sub CMD_INPUT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CMD_INPUT_Click
    End If
End Sub

Private Sub CMD_NEW_Click()
    FGRID.CLEAR
    Form_Load
    TXT_TGL.SetFocus
End Sub

Public Sub J_EDIT_FGRID()
    TXT_CELL.Visible = True
    TXT_CELL.Top = FGRID.CellTop + FGRID.Top
    TXT_CELL.Left = FGRID.CellLeft + FGRID.Left

    TXT_CELL.Text = FGRID.Text
    TXT_CELL.SelStart = 0
    TXT_CELL.SelLength = Len(TXT_CELL.Text)

    TXT_CELL.Visible = True
    TXT_CELL.Height = FGRID.CellHeight
    TXT_CELL.Width = FGRID.CellWidth
    TXT_CELL.SetFocus
End Sub

Private Sub FGRID_DblClick()
    J_EDIT_FGRID
End Sub

Private Sub FGRID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FGRID_DblClick
    ElseIf KeyAscii = 27 Then
        CMD_INPUT.Enabled = True
        CMD_INPUT.SetFocus
    End If
End Sub

Private Sub Form_Load()
    FGRID.TextMatrix(0, 1) = "KD BRG"
    FGRID.TextMatrix(0, 2) = "NAMA BRG"
    FGRID.TextMatrix(0, 3) = "JML JUAL"
    FGRID.TextMatrix(0, 4) = "HRG JUAL"
    FGRID.TextMatrix(0, 5) = "HRG BELI"
    FGRID.TextMatrix(0, 6) = "JML HARGA"
    FGRID.ColWidth(1) = 1000
    FGRID.ColWidth(2) = 4000
    FGRID.ColWidth(3) = 800
    FGRID.ColWidth(4) = 1000
    FGRID.ColWidth(5) = 1000
    FGRID.ColWidth(6) = 1300
    FGRID.Rows = 1
    J_KONEKSI
    TXT_TGL.Text = CLD.Value
End Sub

Public Sub J_KONEKSI()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Module1.con.State = 1 Then con.Close
    con.Open Module1.path_db
    Module1.con.CursorLocation = adUseClient
    
    RS_JUAL.Open "select * from penjualan", con, adOpenKeyset, adLockOptimistic
    RS_MASTER_BRG.Open "select * from MASTER_BRG_JASA", con, adOpenKeyset, adLockOptimistic
    RS_MODAL.Open "select * from KEUANGAN", con, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RS_JUAL
    RS_JUAL.Sort = "TGL_JUAL"
    RS_JUAL.MoveLast
End Sub

Private Sub TXT_CELL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        FGRID.SetFocus
    ElseIf KeyAscii = 13 And FGRID.Col = 1 Then
        STOK_VIEW.Show
        dsa_col = FGRID.Col
        dsa_row = FGRID.Row
    ElseIf KeyAscii = 13 And FGRID.Col = 3 Then
        If TXT_CELL.Text = "" Then
            MsgBox "DATA TIDAK BOLEH KOSONG!"
        ElseIf FGRID.TextMatrix(FGRID.Row, 4) = "" Then
            FGRID.Col = 4
            FGRID.TextMatrix(FGRID.Row, 3) = TXT_CELL.Text
            J_EDIT_FGRID
        Else
            FGRID.TextMatrix(FGRID.Row, 3) = TXT_CELL.Text
            If FGRID.TextMatrix(FGRID.Row, 6) = "" Then
                FGRID.TextMatrix(FGRID.Row, 6) = 0
            End If
            TEMP = FGRID.TextMatrix(FGRID.Row, 6)
            FGRID.TextMatrix(FGRID.Row, 6) = FGRID.TextMatrix(FGRID.Row, 3) * FGRID.TextMatrix(FGRID.Row, 4)
            TXT_TOTAL.Text = FGRID.TextMatrix(FGRID.Row, 6) + Int(TXT_TOTAL.Text) - TEMP
            If FGRID.Rows = (dsa_row + 1) Then
                FGRID.Rows = FGRID.Rows + 1
            End If
            FGRID.Col = 1
            FGRID.Row = FGRID.Row + 1
            J_EDIT_FGRID
        End If
    ElseIf KeyAscii = 13 And FGRID.Col = 4 Then
        If TXT_CELL = "" Then
            MsgBox "DATA TIDAK BOLEH KOSONG!"
        Else
            FGRID.TextMatrix(FGRID.Row, 4) = TXT_CELL.Text
            If FGRID.TextMatrix(FGRID.Row, 6) = "" Then
                FGRID.TextMatrix(FGRID.Row, 6) = 0
            End If
            TEMP = FGRID.TextMatrix(FGRID.Row, 6)
            FGRID.TextMatrix(FGRID.Row, 6) = FGRID.TextMatrix(FGRID.Row, 3) * FGRID.TextMatrix(FGRID.Row, 4)
            TXT_TOTAL.Text = FGRID.TextMatrix(FGRID.Row, 6) + Int(TXT_TOTAL.Text) - TEMP
            If FGRID.Rows = (dsa_row + 1) Then
                FGRID.Rows = FGRID.Rows + 1
            End If
            FGRID.Col = 1
            FGRID.Row = FGRID.Row + 1
            J_EDIT_FGRID
        End If
    ElseIf KeyAscii = 13 And FGRID.Col = 5 Then
        If TXT_CELL = "" Then
            MsgBox "DATA TIDAK BOLEH KOSONG!"
        Else
            FGRID.TextMatrix(FGRID.Row, 5) = TXT_CELL.Text
            If FGRID.TextMatrix(FGRID.Row, 6) = "" Then
                FGRID.TextMatrix(FGRID.Row, 6) = 0
            End If
            TEMP = FGRID.TextMatrix(FGRID.Row, 6)
            FGRID.TextMatrix(FGRID.Row, 6) = FGRID.TextMatrix(FGRID.Row, 3) * FGRID.TextMatrix(FGRID.Row, 4)
            TXT_TOTAL.Text = FGRID.TextMatrix(FGRID.Row, 6) + Int(TXT_TOTAL.Text) - TEMP
            If FGRID.Rows = (dsa_row + 1) Then
                FGRID.Rows = FGRID.Rows + 1
            End If
            FGRID.Col = 1
            FGRID.Row = FGRID.Row + 1
            J_EDIT_FGRID
        End If
    
    ElseIf KeyAscii = 13 And FGRID.Col = 6 Then
        If TXT_CELL = "" Then
            MsgBox "DATA TIDAK BOLEH KOSONG!"
        Else
            If FGRID.TextMatrix(FGRID.Row, 6) = "" Then
                FGRID.TextMatrix(FGRID.Row, 6) = 0
            End If
            TEMP = FGRID.TextMatrix(FGRID.Row, 6)
            FGRID.TextMatrix(FGRID.Row, 6) = TXT_CELL.Text
            FGRID.TextMatrix(FGRID.Row, 4) = FGRID.TextMatrix(FGRID.Row, 6) / FGRID.TextMatrix(FGRID.Row, 3)
            TXT_TOTAL.Text = FGRID.TextMatrix(FGRID.Row, 6) + Int(TXT_TOTAL.Text) - TEMP
            If FGRID.Rows = (dsa_row + 1) Then
                FGRID.Rows = FGRID.Rows + 1
            End If
            FGRID.Col = 1
            FGRID.Row = FGRID.Row + 1
            J_EDIT_FGRID
        End If
    End If
End Sub

Private Sub TXT_CELL_LostFocus()
    TXT_CELL.Visible = False
End Sub

Private Sub TXT_TGL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CLD.Visible = True
        CLD.SetFocus
    End If
End Sub
