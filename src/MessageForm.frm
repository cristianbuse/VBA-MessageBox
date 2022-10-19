VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageForm 
   Caption         =   "UserForm1"
   ClientHeight    =   525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1935
   OleObjectBlob   =   "MessageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event QueryClose(Cancel As Integer, CloseMode As Integer)

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent QueryClose(Cancel, CloseMode)
End Sub

