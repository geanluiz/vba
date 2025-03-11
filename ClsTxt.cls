VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents tb As MSForms.textBox

Sub Init(tbox As Object)
    Set tb = tbox
End Sub

Private Sub tb_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger)

	If InStr(tb.Value, ",") Then
		tb.Value = Replace(tb.Value, ",", ".")
	End If

End Sub

Function getName() As String
	getName = tb.Name
End Function