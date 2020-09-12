VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreviewForm 
   Caption         =   "Технические требования"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10785
   OleObjectBlob   =   "PreviewForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PreviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public abort As Boolean

Private Sub btnCancel_Click()
    abort = True
    Me.Hide
End Sub

Private Sub btnOk_Click()
    abort = False
    Me.Hide
End Sub
