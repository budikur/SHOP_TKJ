VERSION 5.00
Begin VB.Form MAIN_MENU 
   Caption         =   "MENU UTAMA"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   18975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   15240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu LOAD_BD 
      Caption         =   "LOAD DATABASE"
      Index           =   0
   End
   Begin VB.Menu MTR_BRG 
      Caption         =   "MASTER BARANG"
      Index           =   1
   End
   Begin VB.Menu STK_BRG 
      Caption         =   "STOK BARANG"
      Index           =   2
   End
   Begin VB.Menu BELI_BRG 
      Caption         =   "PEMBELIAN BARANG"
      Index           =   3
   End
   Begin VB.Menu JUAL_BRG 
      Caption         =   "PENJUALAN BARANG"
      Index           =   4
   End
   Begin VB.Menu LOST 
      Caption         =   "PENGELUARAN"
      Index           =   5
   End
   Begin VB.Menu LARUG 
      Caption         =   "LABA/RUGI"
      Index           =   6
   End
   Begin VB.Menu KELUAR 
      Caption         =   "KELUAR"
      Index           =   7
   End
End
Attribute VB_Name = "MAIN_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BELI_BRG_Click(Index As Integer)
    PEMBELIAN.Show
End Sub

Private Sub BELI_Click()

End Sub

Private Sub CMD_KELUAR_Click()

End Sub

Private Sub CMD_LABA_RUGI_Click()

End Sub

Private Sub EXIT_Click()

End Sub

Private Sub Form_Load()
    Open App.Path & "\PATH_MDB.txt" For Input As #1
    If EOF(1) Then
        Close #1
        LOAD_Click
    Else
        Line Input #1, Module1.path_db
        Close #1
    End If
    Label1.Caption = Module1.path_db
End Sub

Private Sub JUAL_BRG_Click(Index As Integer)
    PENJUALAN.Show
End Sub

Private Sub LABA_RUGI_Click()
    LABA_RUGI.Show
End Sub

Private Sub LOAD_Click()

End Sub

Private Sub MASTER_BRG_Click(Index As Integer)
    MASTER_BRG.Show
End Sub

Private Sub JUAL_Click()

End Sub

Private Sub KELUAR_Click(Index As Integer)
    End
End Sub

Private Sub LARUG_Click(Index As Integer)
    LABA_RUGI.Show
End Sub

Private Sub LOAD_BD_Click(Index As Integer)
    LOAD_DB.Show
End Sub

Private Sub MASTER_Click()

End Sub

Private Sub LOST_Click(Index As Integer)
    PENGELUARAN.Show
End Sub

Private Sub MTR_BRG_Click(Index As Integer)
    MASTER_BRG.Show
End Sub

Private Sub PENGELUARAN_Click(Index As Integer)
    PENGELUARAN.Show
End Sub

Private Sub STK_BRG_Click(Index As Integer)
    STOK_BRG.Show
End Sub

Private Sub STOK_Click()

End Sub
