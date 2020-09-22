VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Form2"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
FadeIn Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FadeOut Me
End Sub


