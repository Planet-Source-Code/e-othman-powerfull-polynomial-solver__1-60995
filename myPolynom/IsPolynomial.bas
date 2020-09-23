Attribute VB_Name = "IsPolynomial"
Option Explicit

Sub Test_It1()
  Debug.Print "========="
    Debug.Print IsPolynomial1("2*X^3-4*X^2+5*X^1-6") '=True
    Debug.Print IsPolynomial1("2.123*X^3-4.56*X^2+5.8*X^1-6") '=True
    Debug.Print IsPolynomial1("2.123*X-3") '=False, X has no explicitly stated exponent 1
    Debug.Print IsPolynomial1("2.123*X^1-3") '=True
    Debug.Print IsPolynomial1("2.123*X^4") '=True
    Debug.Print IsPolynomial1("2.123*X^4.1") '=False, exponent not integer
    Debug.Print ("+++")
    Debug.Print IsPolynomial1("2.123*x^4") '=True
    Debug.Print IsPolynomial1("2.123*X^4") '=True
    Debug.Print IsPolynomial1("(2.123*X^4)") '=False, parentheses around expression
    Debug.Print IsPolynomial1("") 'Invalid input
  Debug.Print "========="
End Sub

Sub Test_It2()
  Debug.Print "========="
    Debug.Print IsPolynomial2("2*X^3-4*X^2+5*X^1-6") '=True
    Debug.Print IsPolynomial2("2.123*X^4.1") '=False, exponent not integer
    Debug.Print IsPolynomial2("2*Z^4", "Z") '=True
    Debug.Print IsPolynomial2("2*Z^4-Z^3+1*Z", "Z") '=True
    Debug.Print IsPolynomial2("2*X^2-X^3+X") '=True
    Debug.Print IsPolynomial2("2.123*Z^4-Z", "Z") '=True
    Debug.Print ("+++")
    Debug.Print IsPolynomial2("(2.123*Z^4-Z)", "Z") '=False, parentheses around expression
    Debug.Print IsPolynomial2("", "Z") '=True
  Debug.Print "========="
End Sub

Function IsPolynomial1(s As String) As Boolean
'// Dana DeLouis:  dana2@msn.com
'
  Dim RExp As RegExp
  Dim Str As String
  Dim vMatches

  Set RExp = New RegExp
  'Remove Spaces, to make this example 'easier'
  Str = Replace(s, Space(1), vbNullString)

  RExp.Global = True
  RExp.IgnoreCase = True
  RExp.Pattern = "(?:[+-]?\d+\.?\d*[*|/]?X\^\d|[+-]?\d+\.?\d*)*"

  Set vMatches = RExp.Execute(Str)

  If vMatches(1) <> vbNullString Or vMatches.Count > 2 Then
    IsPolynomial1 = False
  Else
    IsPolynomial1 = (vMatches(0) = Str)
  End If
End Function

Function IsPolynomial2(s As String, Optional IndVar As String = "X") As Boolean
'// Dana DeLouis:  dana2@msn.com
'
  Dim RExp As RegExp
  Dim vMatches
  Dim Match
  Dim Str As String
  Dim v(1 To 5) As String
  Dim Temp As String

  Set RExp = New RegExp

  'Remove Spaces, to make this example 'easier'
  Str = Replace(s, Space(1), vbNullString)

  v(1) = "(?:[+-]?\d+\.?\d*\*_X\^\d)"
  v(2) = "(?:[+-]?\d+\.?\d*\*_X)"
  v(3) = "(?:[+-]?_X\^\d)"
  v(4) = "(?:[+-]?_X)"
  v(5) = "(?:[+-]?\d+\.?\d*)"

  With RExp
    .Global = True
    .IgnoreCase = True
    .Pattern = Replace(Join(v, "|"), "_X", IndVar)
    Set vMatches = .Execute(Str)
  End With

  Temp = vbNullString
  For Each Match In vMatches
    Temp = Temp & Match
  Next

  IsPolynomial2 = Temp = Str

  Set RExp = Nothing
End Function
