VERSION 5.00
Begin VB.Form YES_NO 
   Caption         =   "MESSAGE BOX"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton NO 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton YES 
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "RECORD SUDAH ADA, APAKAH DIREPLACE?"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "YES_NO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub NO_Click()
    YES_NO.Hide
    Module1.SORTING
    Module1.DISPLAY
    Module1.DEFAULT
End Sub

Private Sub YES_Click()
    Module1.UPDATE
    YES_NO.Hide
    Module1.SORTING
    Module1.DISPLAY
    Module1.DEFAULT
End Sub
