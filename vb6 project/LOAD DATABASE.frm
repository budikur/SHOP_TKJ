VERSION 5.00
Begin VB.Form LOAD_DB 
   Caption         =   "LOAD DATABASE"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "ADDRESS :"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "LOAD_DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
    File1.FileName = Dir1.Path
    Text1.Text = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dir1_Change
    End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Text1.Text = Dir1.Path
End Sub

Private Sub File1_Click()
    Text1.Text = File1.Path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
    Text1.Text = File1.Path & "\" & File1.FileName
    Module1.path_db = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text1.Text
    
    Open App.Path & "\PATH_MDB.txt" For Output As #1
    Print #1, Module1.path_db
    Close #1
    MAIN_MENU.Label1 = Module1.path_db
    Me.Hide
    MAIN_MENU.Show
End Sub

Private Sub Form_Load()
    Text1.Text = Dir1.Path
End Sub
