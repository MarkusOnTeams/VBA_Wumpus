VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "Deine Aktion?"
   ClientHeight    =   8955.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    End

End Sub

Private Sub OkButton_Click()

    Me.Hide

End Sub


Private Sub UserForm_Activate()

    EingabeTextbox.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True

End Sub

Sub testMsgBox()

MsgBox "Hallo"
End Sub
