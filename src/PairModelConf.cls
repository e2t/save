VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PairModelConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public model As ModelDoc2
Public conf As String

Public Sub Class_Initialize()
    Set model = Nothing
    conf = ""
End Sub
