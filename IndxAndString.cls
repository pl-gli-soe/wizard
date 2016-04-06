VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndxAndString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta w zamierzeniu miala byc pojedynczym itemem jesli nie bede wykorzystywal obiektow
' klasy Dictionary
' kwestia do rozpatrzenia

Public indx As Integer
Public column_name As String

Private Sub Class_Initialize()
    indx = -1
    column_name = ""

End Sub

