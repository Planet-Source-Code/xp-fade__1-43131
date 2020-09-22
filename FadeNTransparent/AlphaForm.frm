VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identification"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   ClipControls    =   0   'False
   Icon            =   "AlphaForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Enabled         =   0   'False
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accepter"
      Height          =   405
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Mot de passe:"
      Height          =   225
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
Main.Show

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
FadeIn Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FadeOut Me
End Sub

