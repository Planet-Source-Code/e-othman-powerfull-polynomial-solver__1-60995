Attribute VB_Name = "Mod_Poly"
Option Explicit
Public P() As Variant

'Private oEqu As clsEquation



Public Sub P_Read(oEqu As clsEquation)
Dim i As Integer
Dim N As Integer

oEqu.Parse_Poly
N = oEqu.Deg_Poly
ReDim P(1 To N)

For i = 1 To N
If i > oEqu.Pow.Count Then
P(i) = 0
Else
P(i) = oEqu.Coeff(i)
End If

Next i
End Sub
