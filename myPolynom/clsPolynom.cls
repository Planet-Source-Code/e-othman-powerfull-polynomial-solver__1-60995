VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolynom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    

Private Dirty               As Boolean
Private P_Error             As Integer
Public flag                 As Integer
Private isRat               As Boolean
Public dRoots               As Collection
Public P                    As Collection
Public Result_p             As Collection
Public Result_q             As Collection

Public F_p                  As Collection
Public F_q                  As Collection

Public F_pPlusq              As Collection
Public F_pMinusq             As Collection
Public F_pByq                As Collection
Public R_p                   As Collection
Public R_q                   As Collection
Public R_pByq               As Collection
Private Parsed              As Boolean
Private Deg                 As Boolean
Private Equ                 As String
Public dAnswer              As Double
Private EquParsed           As Collection
Private EquOrder            As Collection
Private Vars                As Collection
Public Coeff                As Collection
Public Pow                  As Collection
Public Msg                  As Collection
Public unkown               As String
Private Const PRI_ADD       As Integer = 1
Private Const PRI_MOD       As Integer = 2
Private Const PRI_MUL       As Integer = 3
Private Const PRI_NEG       As Integer = 4
Private Const PRI_EXP       As Integer = 5
Private Const PRI_VAR       As Integer = 6
Private Const PRI_PAR       As Integer = 7
Private Const PRI_LEVEL     As Integer = 7

Private Const EQ_NONE       As Integer = 0
Private Const EQ_STRING     As Integer = 1
Private Const EQ_NUMBER     As Integer = 2

Private Const ER_NONE       As Integer = 0
Private Const ER_VAR        As Integer = 1

Private Const EQ_PAREN      As Integer = 1100
Private Const MaxDeg        As Long = 99999
Private Const EQ_FUNCTION   As Integer = 1101
Private Const EQ_VARIABLE   As Integer = 1102
Private Const EQ_INVALID    As Integer = 1103
Private Const EQ_ARGS       As Integer = 1104
Private Const EQ_NAME       As Integer = 1105
Const MAXR As Double = 1E+15
 Const MINR As Double = 0.00000001
 Const MAXG As Double = 10
 Const MAXI As Double = 50
 Const prfMaxIterations As Integer = 2000
 Const prfTolerance As Double = 0.0000000001
 Const prfInitialGuess As Double = 0.005
'Declare matrix type

Private Const PI            As Double = 3.14159265358979
Private Const DEG_TO_RAD    As Double = PI / 180    ''0.01745329251995
Private Const RAD_TO_DEG    As Double = 180 / PI    ''57.2957795131

Private Sub Class_Initialize()
    
    Set EquParsed = New Collection
    Set EquOrder = New Collection
    Set Vars = New Collection
    
    Dirty = False
    Parsed = True
    Degrees = False
    
End Sub

Private Sub Class_Terminate()

    Set EquParsed = Nothing
    Set EquOrder = Nothing
    Set Vars = Nothing

End Sub


Public Property Let Degrees(B As Boolean)
    If B <> Deg Then
        Deg = B
        Dirty = True
    End If
End Property

Public Property Get Degrees() As Boolean
    Degrees = Deg
End Property

Public Property Let Equation(e As String)
    Parsed = False
    Dirty = True
    Equ = LCase(e)
End Property

Public Property Get Equation() As String
    Equation = Equ
End Property

Public Property Get Var(Name As String) As Double
    On Local Error GoTo GetError
    Var = Round(Val(Vars(Name)))
    Exit Property
GetError:
    Var = 0#
End Property

Public Property Let Var(Name As String, Num As Double)
    On Local Error Resume Next
    Dirty = True
    Vars.Remove Trim$(Name)
    Vars.Add Num, Trim$(Name)
End Property

Public Sub VarClear()
    Set Vars = New Collection
    Dirty = True
End Sub

Public Sub VarRemove(Name As String)
    On Local Error Resume Next
    Vars.Remove Name
    Dirty = True
End Sub

Public Function Solution() As Double
    If Dirty Then
        Solve
    End If
    Solution = dAnswer
End Function

Public Sub Solve()
    Dim i                   As Long
    Dim j                   As Long
    Dim L                   As Long
    Dim r                   As Long
    Dim m                   As Long
    Dim N                   As Long
    Dim X                   As Double
    Dim Y                   As Double
    Dim v                   As Variant
    Dim eSpace              As Integer
    Dim Temp()              As Variant
'    Dim f                   As clsEquation
    Dim j2                  As Long
    
    On Local Error GoTo SolveError
    
    If Not Parsed Then
        Parse
    End If
    ReDim Temp(1 To EquParsed.Count)
    For i = 1 To EquParsed.Count
        Temp(i) = EquParsed(i)
    Next
    eSpace = ER_NONE
    For i = 1 To EquOrder.Count
        m = EquOrder(i)
        v = Temp(m)
        Select Case v
            Case "~"
                r = GetRight(m, Temp)
                Temp(m) = -Round(Val(Temp(r)))
                Temp(r) = Null
            Case "*"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) * Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "/"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) / Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "\"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) \ Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "+"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) + Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "-"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) - Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "^"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) ^ Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "%"
                L = GetLeft(m, Temp)
                r = GetRight(m, Temp)
                Temp(L) = Round(Val(Temp(L))) Mod Round(Val(Temp(r)))
                Temp(r) = Null
                Temp(m) = Null
            Case "("
                i = i + 1
                N = EquOrder(i)
                r = GetRight(m, Temp)
                If r >= N Then
                    Temp(m) = 0#
                    Temp(N) = Null
                Else
                    Temp(m) = Temp(r)
                    Temp(r) = Null
                    Temp(N) = Null
                End If
            Case Else
                If Right$(Temp(m), 1) = "(" Then
                    i = i + 1
                    N = EquOrder(i)
                    L = GetRight(m, Temp)
                    r = GetLeft(N, Temp)
                    If L >= N Then
                        Err.Raise EQ_ARGS, "clsEquation", "Invalid arguments to function: " & v & ")"
                        Exit Sub
                    Else
                        X = Round(Val(Temp(L)))
                    End If
                    If r <= m Then
                        Err.Raise EQ_ARGS, "clsEquation", "Invalid arguments to function: " & v & ")"
                        Exit Sub
                    Else
                        Y = Round(Val(Temp(r)))
                    End If
                    Temp(r) = Null
                    Temp(L) = Null
                    Temp(m) = Null
                    Temp(N) = Null
                    Select Case v
                        Case "abs("
                            Temp(m) = Abs(X)
                        Case "atn("
                            If Degrees Then
                                Temp(m) = Atn(X) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X)
                            End If
                        Case "arctan("
                            If Degrees Then
                                Temp(m) = Atn(X) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X)
                            End If
                        Case "cos("
                            If Degrees Then
                                Temp(m) = Cos(X * DEG_TO_RAD)
                            Else
                                Temp(m) = Cos(X)
                            End If
                        Case "exp("
                            Temp(m) = Exp(X)
                        Case "fix("
                            Temp(m) = Fix(X)
                        Case "int("
                            Temp(m) = Int(X)
                        Case "log("
                            Temp(m) = Log(X)
                        Case "rnd("
                            Temp(m) = Rnd(X)
                        Case "sgn("
                            Temp(m) = Sgn(X)
                        Case "sin("
                            If Degrees Then
                                Temp(m) = Sin(X * DEG_TO_RAD)
                            Else
                                Temp(m) = Sin(X)
                            End If
                        Case "sqr("
                            Temp(m) = Sqr(X)
                        Case "tan("
                            If Degrees Then
                                Temp(m) = Tan(X * DEG_TO_RAD)
                            Else
                                Temp(m) = Tan(X)
                            End If
                        Case "min("
                            Temp(m) = IIf(X < Y, X, Y)
                        Case "max("
                            Temp(m) = IIf(X > Y, X, Y)
                        Case "random("
                            Temp(m) = (Rnd * (Y - X)) + X
                        Case "mod("
                            Temp(m) = X Mod Y
                        Case "logn("
                            Temp(m) = Log(X) / Log(Y)
                        Case "rand("
                            Temp(m) = Int(Rnd * X)
                        Case "sec("
                            If Degrees Then
                                Temp(m) = (1 / Cos(X * DEG_TO_RAD))
                            Else
                                Temp(m) = 1 / Cos(X)
                            End If
                        Case "cosec("
                            If Degrees Then
                                Temp(m) = (1 / Sin(X * DEG_TO_RAD))
                            Else
                                Temp(m) = 1 / Sin(X)
                            End If
                        Case "cotan("
                            If Degrees Then
                                Temp(m) = (1 / Tan(X * DEG_TO_RAD))
                            Else
                                Temp(m) = 1 / Tan(X)
                            End If
                        Case "arcsin("
                            If Degrees Then
                                Temp(m) = (Atn(X / Sqr(-X * X + 1))) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X / Sqr(-X * X + 1))
                            End If
                        Case "arccos("
                            If Degrees Then
                                Temp(m) = (Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
                            End If
                        Case "arcsec("
                            If Degrees Then
                                Temp(m) = (Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
                            End If
                        Case "arccosec("
                            If Degrees Then
                                Temp(m) = (Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
                            End If
                        Case "arccotan("
                            If Degrees Then
                                Temp(m) = (Atn(X * DEG_TO_RAD) + 2 * Atn(1)) * RAD_TO_DEG
                            Else
                                Temp(m) = Atn(X) + 2 * Atn(1)
                            End If
                        Case "sinh("
                            Temp(m) = (Exp(X) - Exp(-X)) / 2
                        Case "cosh("
                            Temp(m) = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
                        Case "tanh("
                            Temp(m) = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
                        Case "sech("
                            Temp(m) = 2 / (Exp(X) + Exp(-X))
                        Case "cosech("
                            Temp(m) = 2 / (Exp(X) - Exp(-X))
                        Case "cotanh("
                            Temp(m) = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
                        Case "arcsinh("
                            Temp(m) = Log(X + Sqr(X * X + 1))
                        Case "arccosh("
                            Temp(m) = Log(X + Sqr(X * X - 1))
                        Case "arctanh("
                            Temp(m) = Log((1 + X) / (1 - X)) / 2
                        Case "arcsech("
                            Temp(m) = Log((Sqr(-X * X + 1) + 1) / X)
                        Case "arccosech("
                            Temp(m) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
                        Case "arccotanh("
                            Temp(m) = Log((X + 1) / (X - 1)) / 2
                        Case "log10("
                            Temp(m) = Log(X) / Log(10)
                        Case "log2("
                            Temp(m) = Log(X) / Log(2)
                        Case "ln("
                            Temp(m) = Log(X)
                        Case "deg("
                            Temp(m) = X * RAD_TO_DEG
                        Case "rad("
                            Temp(m) = X * DEG_TO_RAD
                        Case Else
                            Err.Raise EQ_FUNCTION, "clsEquation", "Undefined Function: " & v & ")"
                            Exit Sub
                    End Select
                Else
                    Select Case v
                        Case "pi":
                            Temp(m) = PI
                        Case "e":
                            Temp(m) = 2.718281828
                        Case "rnd":
                            Temp(m) = Rnd
                        Case Else
                            eSpace = ER_VAR
                            Temp(m) = Round(Val(Vars(Temp(m))))
                            eSpace = ER_NONE
                    End Select
                End If
        End Select
    Next i
   
    dAnswer = Trim(Round(Val(Temp(GetRight(0, Temp)))))
    Dirty = False
    Exit Sub

SolveError:
    Select Case Err
        Case 6, 11, EQ_PAREN To EQ_NAME
            Err.Raise Err, "clsEquation", Err.Description
        Case 5:
            Select Case eSpace
                Case ER_VAR
                    Err.Raise EQ_VARIABLE, "clsEquation", "Undefined Variable: " & v
                Case Else
                    Err.Raise Err, "clsEquation", Err.Description
            End Select
        Case Else
            Err.Raise EQ_INVALID, "clsEquation", "Invalid Equation"
    End Select
End Sub

Private Sub Parse()
    Dim i                   As Integer
    Dim s                   As String
    Dim t                   As Integer
    Dim j                   As Integer
    Dim sTmp                As String
    Dim P                   As Integer
    Dim EquPriority         As Collection
    Dim maxPriority         As Integer
    Dim isNeg               As Boolean
    Dim k As Integer
    On Local Error GoTo Err_Parse
    
    s = ""
    t = EQ_NONE
    j = 1
    P = 0
    k = 1
    isNeg = False
    Set EquPriority = New Collection
    Set EquParsed = New Collection
   
    EquParsed.Add ""
    EquPriority.Add ""
    
    maxPriority = PRI_LEVEL
    
    ' added to support 'mod' in addition to %
    Equ = Replace(Equ, " mod ", "%")
    
    For i = 1 To Len(Equ)
        sTmp = Mid$(Equ, i, 1)
        Select Case sTmp
            Case "A" To "Z", "a" To "z", "_"
            If t = EQ_NONE Then
                t = EQ_STRING
                s = sTmp
                
            ElseIf t = EQ_NUMBER Then
                t = EQ_STRING
                EquParsed.Add s, , j
                EquPriority.Add 0, , j
                
               
                j = j + 1
                EquParsed.Add "*", , j
                EquPriority.Add PRI_MUL + P, , j
                j = j + 1
                s = sTmp
            Else
                s = s + sTmp
                'varIndex.Add Str(j), j
            End If
            isNeg = True
      Case "1" To "9", "0", "."
            If t = EQ_NONE Then
                t = EQ_NUMBER
                s = sTmp
            Else
                s = s + sTmp
            End If
            isNeg = True
      Case "(":
            If t = EQ_STRING Then
                EquParsed.Add s + sTmp, , j
                EquPriority.Add P + PRI_PAR, , j
                j = j + 1
                s = ""
            ElseIf t = EQ_NUMBER Then
                EquParsed.Add s, , j
                EquPriority.Add 0, , j
                j = j + 1
                EquParsed.Add "*", , j
                EquPriority.Add P + PRI_MUL, , j
                j = j + 1
                EquParsed.Add sTmp, , j
                EquPriority.Add P + PRI_PAR, , j
                j = j + 1
                s = ""
            Else
                EquParsed.Add sTmp, , j
                EquPriority.Add P + PRI_PAR, , j
                j = j + 1
            End If
            P = P + PRI_LEVEL
            t = EQ_NONE
            If maxPriority < P + PRI_LEVEL Then
                maxPriority = P + PRI_LEVEL
            End If
            isNeg = False
        Case "*", "/":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add P + PRI_MUL, , j
            j = j + 1
            t = EQ_NONE
            isNeg = False
        Case "\":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add P + PRI_MUL, , j
            j = j + 1
            t = EQ_NONE
            isNeg = False
        Case "+":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
                EquParsed.Add sTmp, , j
                EquPriority.Add P + PRI_ADD, , j
                j = j + 1
                t = EQ_NONE
            Else
                ' added by dd
                EquParsed.Add sTmp, , j
                EquPriority.Add P + PRI_ADD, , j
                j = j + 1
                t = EQ_NONE
            End If
            isNeg = False
        Case "-":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            If isNeg Then
                EquParsed.Add sTmp, , j
                EquPriority.Add P + PRI_ADD, , j
                j = j + 1
                t = EQ_NONE
            Else
                EquParsed.Add "~", , j
                EquPriority.Add P + PRI_NEG, , j
                j = j + 1
                t = EQ_NONE
            End If
            isNeg = False
        Case "^":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add P + PRI_EXP, , j
            j = j + 1
            t = EQ_NONE
            isNeg = False
        Case "%":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add P + PRI_MOD, , j
            j = j + 1
            t = EQ_NONE
            isNeg = False
        Case ",":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add Null, , j
            EquPriority.Add 0, , j
            j = j + 1
            t = EQ_NONE
            isNeg = False
        Case ")":
            If t <> EQ_NONE Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add P - (PRI_LEVEL - PRI_PAR), , j
            P = P - PRI_LEVEL
            j = j + 1
            t = EQ_NONE
            isNeg = True
        End Select
    Next i
    If s <> "" Then
        EquParsed.Add s, , j
        EquPriority.Add IIf(t = EQ_STRING, P + PRI_VAR, 0), , j
        j = j + 1
    End If
    EquParsed.Remove j
    EquPriority.Remove j
    If P <> 0 Then
        Err.Raise EQ_PAREN, "clsEquation", "Unbalanced parenthesis"
        Set EquPriority = Nothing
        Exit Sub
    End If
    Set EquOrder = New Collection
    EquOrder.Add ""
    For j = 1 To maxPriority
        For i = EquPriority.Count To 1 Step -1
            If EquPriority(i) = j Then
                EquOrder.Add i, , , 1
            End If
        Next i
    Next j
    EquOrder.Remove 1
    Parsed = True
    Set EquPriority = Nothing
    
    Exit Sub

Err_Parse:
    Set EquPriority = Nothing
    Err.Raise Err, "clsEquation.Parse", Err.Description
    
End Sub

Private Function GetRight(ByVal j As Long, v() As Variant) As Long
    ' helper function for solve
    Dim i As Long
    For i = j + 1 To UBound(v)
        If Not IsNull(v(i)) Then
            GetRight = i
            Exit Function
        End If
    Next i
    GetRight = 0
End Function

Private Function GetLeft(ByVal j As Long, v() As Variant) As Long
    ' helper function for solve
    Dim i As Long
    For i = j - 1 To 1 Step -1
        If Not IsNull(v(i)) Then
            GetLeft = i
            Exit Function
        End If
    Next i
    GetLeft = 0
End Function


Public Sub Parse_Poly()
Dim N As Long
Dim i As Long
Dim v As Variant
Dim j As Long
Set Coeff = New Collection
Set Pow = New Collection
Set Msg = New Collection
Dim Temp() As Variant
Dim L As Long
Dim r As Long
Dim coef As Double
Dim coef0 As Double
Dim power As Long
If Not Parsed Then
        Parse
    End If
    Msg.Add ""
    j = 1
    coef0 = 0
 
 ReDim Temp(1 To EquParsed.Count)
    For i = 1 To EquParsed.Count
        Temp(i) = EquParsed(i)
    Next
    N = EquParsed.Count
For i = 1 To EquParsed.Count
 v = Temp(i)
Select Case v
            Case "A" To "Z", "a" To "z", "_"
            unkown = Temp(i)
            If N = 1 Then
            power = 1
            coef = 1
            Coeff.Add coef
            Pow.Add power
            End If
            If i = 1 And N > 1 Then
            coef = 1
            r = GetRight(i, Temp)
            If Temp(r) = "^" Then
            r = GetRight(r, Temp)
            power = Val(Temp(r))
            Coeff.Add coef
            Pow.Add power
            End If
            r = GetRight(i, Temp)
            If Temp(r) <> "^" Then
            power = 1
            coef = 1
            Coeff.Add coef
            Pow.Add power
            
            End If
            
            End If
            If i = 2 Then
            r = GetRight(i, Temp)
            L = GetLeft(i, Temp)
            If Temp(r) = "^" Then
            r = GetRight(r, Temp)
            power = Val(Temp(r))
            coef = 1
            'l = GetLeft(l, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            
            End If
            Coeff.Add coef
            Pow.Add power
            
            
            End If
             r = GetRight(i, Temp)
            L = GetLeft(i, Temp)
            If Temp(r) <> "^" Then
            power = 1
            coef = 1
            'l = GetLeft(l, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            
            End If
            Coeff.Add coef
            Pow.Add power
            End If
            End If
            
            If i < EquParsed.Count - 1 And i > 2 Then
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(L) = "*" And Temp(r) = "^" Then
            
            L = GetLeft(L, Temp)
            r = GetRight(r, Temp)
            
            coef = Val(Temp(L)) 'get the coefficient
            power = Val(Temp(r))
            If L > 1 Then
            L = GetLeft(L, Temp)
            
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            End If
            End If
            Coeff.Add coef
            Pow.Add power
            End If
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(L) <> "*" And Temp(r) = "^" Then
            r = GetRight(r, Temp)
            
            coef = 1
            power = Val(Temp(r))
            
            'If l > 1 Then
            'l = GetLeft(l, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            'End If
            End If
            Coeff.Add coef
            Pow.Add power
            End If
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(r) <> "^" And Temp(L) = "*" Then
            L = GetLeft(L, Temp)
            coef = Val(Temp(L))
            power = 1
            If L > 1 Then
            L = GetLeft(L, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            End If
            End If
            Coeff.Add coef
            Pow.Add power
            End If
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(r) <> "^" And Temp(L) <> "*" Then
            power = 1
            coef = 1
            'If l > 1 Then
            'l = GetLeft(l, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            'End If
            End If
            Coeff.Add coef
            Pow.Add power
            End If
            End If
            'l = GetLeft(i, Temp)
            'r = GetRight(i, Temp)
            If i = EquParsed.Count And N > 1 Then
            L = GetLeft(i, Temp)
            If Temp(L) = "*" Then
            
            L = GetLeft(L, Temp)
            
            coef = Val(Temp(L)) 'get the coefficient
            power = Val(1)
           
            If L > 1 Then
            L = GetLeft(L, Temp)
            
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            End If
            End If
            Coeff.Add coef
            Pow.Add power
             End If
             L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(L) <> "*" Then
            coef = Val(1)
            power = Val(1)
            
            'If l > 1 Then
            'l = GetLeft(l, Temp)
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef = -coef
            End If
            'End If
            Coeff.Add coef
            Pow.Add power
            End If
            End If
            'End If
            'Coeff.Add coef
            'Pow.Add power
            'End If
            
            Case "1" To CStr(MaxDeg), "0", "."
            If N = 1 Then
            power = 0
            coef0 = Val(Temp(N))
           
            
            End If
            If N = 2 Then
            power = 0
            L = GetLeft(i, Temp)
            coef0 = -Val(Temp(N))
            
            End If
            If i < EquParsed.Count And i > 2 And N > 2 Then
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            If Temp(r) <> "*" And Temp(L) <> "^" Then
             
             'If l > 1 Then
             'l = GetLeft(l, Temp)
            
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef0 = coef0 - Val(Temp(i))
            Else
            coef0 = coef0 + Val(Temp(i))
            End If
            'End If
            'Coeff.Add coef0
            End If
            End If
            If i = EquParsed.Count And N > 2 Then
            
            L = GetLeft(i, Temp)
            If Temp(L) <> "^" Then
            
            
             If Temp(L) = "-" Or Temp(L) = "~" Then
            
            coef0 = coef0 - Val(Temp(i))
            Else
            coef0 = coef0 + Val(Temp(i))
            End If
            End If
            End If
            If i = 1 And N > 2 Then
            
            r = GetRight(i, Temp)
            If Temp(r) <> "*" Then
            coef0 = coef0 + Val(Temp(i))
            'Coeff.Add coef0
            End If
            End If
            
            If i = 2 And N > 2 Then
            L = GetLeft(i, Temp)
            r = GetRight(i, Temp)
            'If l >= 1 Then
            If Temp(r) <> "*" And Temp(L) <> "^" Then
            ' coef0 = coef0 + Val(Temp(i))
             'If l > 1 Then
            ' l = GetLeft(l, Temp)
            
            If Temp(L) = "-" Or Temp(L) = "~" Then
            coef0 = coef0 - Val(Temp(i))
            Else
            coef0 = coef0 + Val(Temp(i))
           ' End If
            
            'Coeff.Add coef0
            
            
            End If
            
            End If
            
            End If
            
            
            Case Else
            
            End Select
            
Next i
Coeff.Add coef0
Pow.Add 0
Resolve_Cf
Sort

'msg.Add Chr(13) + Chr(10) + "Unkown : " & unkown
'msg.Add Chr(13) + Chr(10) + "Coefficients : " + Chr(13) + Chr(10)
'For i = 1 To Coeff.Count

'msg.Add Chr(13) + Chr(10) + "a" & Str(Pow(i)) & " = " & Str(Coeff(i))
'Next i

'MsgBox msg, vbOKOnly

End Sub


Private Sub Resolve_Cf()
Dim ptmp() As Variant
Dim CTmp() As Variant
Dim i As Long
Dim j As Long
Dim Plen As Long
Dim Clen As Long

If Not Parsed Then
        Parse
    End If
    
    ReDim ptmp(1 To Pow.Count)
    ReDim CTmp(1 To Coeff.Count)
    Plen = Pow.Count
    Clen = Coeff.Count
    
 For i = 1 To Pow.Count
 ptmp(i) = Pow(i)
 Next
 
 For i = 1 To Coeff.Count
 CTmp(i) = Coeff(i)
 Next
 
 
 For i = 1 To Pow.Count - 1
    For j = i + 1 To Pow.Count
        If ptmp(i) = ptmp(j) Then
        CTmp(i) = CTmp(i) + CTmp(j)
        CTmp(j) = Null
        ptmp(j) = Null
        End If
        Next j
        Next i
        
 Set Coeff = New Collection
 Set Pow = New Collection
 
        For i = 1 To Clen
        If Not IsNull(CTmp(i)) Then
        Coeff.Add CTmp(i)
        End If
        Next
        
        For i = 1 To Plen
        If Not IsNull(ptmp(i)) Then
        Pow.Add ptmp(i)
        End If
        Next
        
    
 
    
End Sub

Public Function Deg_Poly(P As Collection)
Dim i As Long
'Dim max As Long

   'max = Pow(1)
   'For i = 1 To Pow.Count
   'If Pow(i) > max Then
  ' max = Pow(i)
  ' End If
  ' Next
   
   Deg_Poly = P.Count - 1
   
   
   



End Function

Public Sub Report()
Dim DG As Long
Dim i As Long

DG = Deg_Poly(P)
Msg.Add "Unkown : " & unkown + Chr(13) + Chr(10)
Msg.Add "Degree :" & Str(DG) + Chr(13) + Chr(10)
For i = 1 To EquParsed.Count
Msg.Add EquParsed(i)
Next
Msg.Add "Coefficients : " + Chr(13) + Chr(10)
For i = P.Count To 1 Step -1
Msg.Add "a" & Str(i - 1) & " = " & Str(P(i)) + Chr(13) + Chr(10)

Next i
Msg.Add CStr("P( 1)= " & P_Value(P, 1)) + Chr(13) + Chr(10)

Msg.Add CStr("P(-1)= " & P_Value(P, -1))
        
End Sub


Public Sub P_Read()
Dim i As Long
Dim N As Long
Dim j As Long
Dim k As Long

Dim tmp() As Variant
Initialize
Set P = New Collection
Parse_Poly
N = Deg_Poly(P)
ReDim tmp(1 To Coeff.Count)

For i = 1 To Coeff.Count
tmp(i) = Coeff(i)
Next i


For i = 1 To Pow.Count - 1
P.Add tmp(i)
j = Pow(i + 1) - Pow(i)
If j > 1 Then
For k = 1 To j - 1

P.Add 0
Next
End If


Next
P.Add tmp(Pow.Count)

End Sub

Public Sub Sort()
Dim ptmp() As Variant
Dim CTmp() As Variant
Dim i As Long
Dim j As Long
Dim Plen As Long
Dim Clen As Long
Dim Inn As Long
Dim out As Long
Dim Temp As Variant


If Not Parsed Then
        Parse
    End If
    
    ReDim ptmp(1 To Pow.Count)
    ReDim CTmp(1 To Coeff.Count)
    Plen = Pow.Count
    Clen = Coeff.Count
    
 For i = 1 To Pow.Count
 ptmp(i) = Pow(i)
 Next
 
 For i = 1 To Coeff.Count
 CTmp(i) = Coeff(i)
 Next
 
 'bubble sort for powers and coefficients
 For out = Plen - 1 To 1 Step -1
 For Inn = 1 To out
 If ptmp(Inn) > ptmp(Inn + 1) Then
 Temp = ptmp(Inn)
 ptmp(Inn) = ptmp(Inn + 1)
 ptmp(Inn + 1) = Temp
 'same for coefficients
 Temp = CTmp(Inn)
 CTmp(Inn) = CTmp(Inn + 1)
 CTmp(Inn + 1) = Temp
 
 End If
 Next
 Next
 
 Set Coeff = New Collection
 Set Pow = New Collection
 
        For i = 1 To Clen
        If Not IsNull(CTmp(i)) Then
        Coeff.Add CTmp(i)
        End If
        Next
        
        For i = 1 To Plen
        If Not IsNull(ptmp(i)) Then
        Pow.Add ptmp(i)
        End If
        Next
        
End Sub

Public Function P_Value(A As Collection, X As Double) As Double
' *** Calculate A(x) for x double. Horner Method***
Dim i As Long
Dim v As Double

v = 0
For i = A.Count To 1 Step -1

v = X * v + A(i)


Next
P_Value = v




End Function

Public Sub Factor(X As Long, F As Collection)
Dim i As Long
Set F = New Collection
For i = 1 To Abs(X)

If X Mod i = 0 Then
F.Add i
F.Add -i
End If
Next



End Sub

Public Sub Rational(P As Collection, R_p As Collection, R_q As Collection)
Dim N As Long
Dim i As Long
Dim j As Long
Dim L As Long
Dim m As Long
Dim k As Long

N = Deg_Poly(P)
Msg.Add "Investigating Possible Rational Roots.First Method.."

Set R_p = New Collection
Set R_q = New Collection
Set F_pByq = New Collection

Factor P(1), F_p
Factor P(N + 1), F_q
Factor P_Value(P, 1), F_pMinusq
Factor P_Value(P, -1), F_pPlusq
For i = 1 To F_p.Count
    For j = 1 To F_q.Count
    F_pByq.Add F_p(i) / F_q(j)
    Next
    Next
    Resolve_Fac F_pByq
    

For i = 1 To F_p.Count
    For j = 1 To F_q.Count
    If (Found((F_p(i) - F_q(j)), F_pMinusq) = True) And (Found((F_p(i) + F_q(j)), F_pPlusq) = True) Then
    
        'For l = 1 To F_pMinusq.Count
        'If (F_p(i) - F_q(j)) = F_pMinusq(l) Then
        'For m = 1 To F_pPlusq.Count
        ' If (F_p(i) + F_q(j)) = F_pPlusq(m) Then
        ' For k = 1 To F_pByq.Count
        ' If (F_p(i) / F_q(j)) = F_pByq(k) Then
        R_p.Add F_p(i)
        R_q.Add F_q(j)
       ' End If
       ' Next
        End If
'Next
'End If

'Next
Next
Next



Resolve_Fac R_p
Resolve_Fac R_q


Final_R
Resolve_R R_p, R_q







End Sub

Public Sub Rational_root(P As Collection)
Rational P, R_p, R_q
Dim i As Long
Dim Str As String
If R_p.Count > 0 And R_q.Count > 0 Then

Msg.Add "If p/q is a rational root for P(" + UCase(unkown) + ")" + Chr(13) + Chr(10)
Msg.Add "Possible solutions are: " + Chr(13) + Chr(10)

Str = "p = " & "{ "
For i = 1 To R_p.Count
Str = Str + CStr(R_p(i)) + " , "
Next
Str = Left(Str, Len(Str) - 3)
Str = Str + "  }" + Chr(13) + Chr(10)
Str = Str + Chr(13) + Chr(10) + "q = " & "{ "
For i = 1 To R_q.Count
Str = Str + CStr(R_q(i)) + " , "
Next
Str = Left(Str, Len(Str) - 3)
Str = Str + "  }" + Chr(13) + Chr(10)
Msg.Add Str

End If

End Sub

Public Sub Resolve_Fac(F As Collection)
 Dim iSelItem  As Long
 Dim iTempItem  As Long

 For iSelItem = F.Count To 1 Step -1
 For iTempItem = 1 To F.Count
 If iTempItem = F.Count Then
 Exit For
 Else
 If F(iTempItem) = F(iSelItem) Then
 If iTempItem <> iSelItem Then
 F.Remove (iSelItem)
 Exit For
 End If
 End If
 End If
 Next iTempItem
 Next iSelItem




End Sub

Public Function Found(Item As Variant, C As Collection) As Boolean

Dim fnd As Boolean
Dim i As Long
fnd = False


For i = 1 To C.Count
If C(i) = Item Then
fnd = True
Exit For
End If

Next

Found = fnd


End Function

Public Sub Final_R()
Dim i As Long
Dim j As Long
Dim ptmp() As Variant
Dim qTmp() As Variant
 If R_p.Count <= 1 Or R_q.Count <= 1 Then
 Exit Sub
 
 End If


If R_p.Count > 1 And R_q.Count > 1 Then
    ReDim ptmp(1 To R_p.Count)
    ReDim qTmp(1 To R_q.Count)
    
    
 For i = 1 To R_p.Count
 ptmp(i) = R_p(i)
 Next
 
 For i = 1 To R_q.Count
 qTmp(i) = R_q(i)
 Next
 Set R_p = New Collection
 Set R_q = New Collection
 
 For i = 1 To UBound(ptmp)
 For j = 1 To UBound(qTmp)
If (Found((ptmp(i) - qTmp(j)), F_pMinusq) = True) And (Found((ptmp(i) + qTmp(j)), F_pPlusq) = True) And (Found((ptmp(i) / qTmp(j)), F_pByq) = True) Then
R_p.Add ptmp(i)
R_q.Add qTmp(j)

End If
 Next
 Next
End If


End Sub

Public Sub Resolve_R(R_p As Collection, R_q As Collection)
Dim iSelItem  As Long
 Dim iTempItem  As Long

 For iSelItem = R_p.Count To 1 Step -1
 For iTempItem = 1 To R_p.Count
 If iTempItem = R_p.Count Then
 Exit For
 Else
 If (R_p(iTempItem) / R_q(iTempItem)) = (R_p(iSelItem) / R_q(iSelItem)) Then
 If iTempItem <> iSelItem Then
 R_p.Remove (iSelItem)
R_q.Remove (iSelItem)
 Exit For
 End If
 End If
 End If
 Next iTempItem
 Next iSelItem



End Sub

Public Function IntCheck(P As Collection) As Boolean
Dim isInteger As Boolean
Dim i As Long
Dim Val As Variant
isInteger = True

For i = 1 To P.Count
 Val = P(i)
 If Val - (Int(Val)) <> 0 Then
  isInteger = False
  Exit For
  
 End If
Next

IntCheck = isInteger

End Function

Public Sub Strategy_Launch()
Dim i As Long
Dim Y As Collection
Dim q As Collection
Dim r As Collection
Set Y = New Collection
Set q = New Collection
Set r = New Collection

Dim N As Long
Initialize
N = Deg_Poly(P)
 
If N = 1 Then
P_Linear P

Exit Sub

End If

If N = 2 Then
P_Quadratic P

Exit Sub

End If

If N > 2 And IntCheck(P) = False Then
Msg.Add "Found Real Coefficients." + Chr(13) + Chr(10)
Msg.Add "Rational roots Method Is not suitable.You May try other Methods.Choose Method from List Below :"
        
Msg.Add "(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu"
flag = 1
Exit Sub

End If

If N > 2 And IntCheck(P) = True Then
If P(1) * P(N + 1) <> 0 Then
If (P_Value(P, 1) * P_Value(P, -1) <> 0) Then


Solve_Rational P
 If Result_p.Count = 0 And Result_q.Count = 0 Then
        
        Msg.Add "this Method failed.You May try other Methods.Choose Method from List Below :"
        
        Msg.Add "(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu"
        flag = 1
       
        End If



Exit Sub
'End If

Else
Solve_RationalBis P
If Result_p.Count = 0 And Result_q.Count = 0 Then
        
        Msg.Add "this Method failed.You May try other Methods.Choose Method from List Below :"
        
        Msg.Add "(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu"
        flag = 1
       
        End If
Exit Sub
End If
End If
End If


If (N > 2) And IntCheck(P) = True Then
If (P(1) = 0 Or P(N + 1) = 0) Then
Y.Add 0
Y.Add 1
P_Euclid P, Y, q, r
Msg.Add " 0 is a trivial root ."
Msg.Add "Reduced Polynomials: " + P_Write(q)
If (P_Value(P, 1) * P_Value(P, -1) = 0) Then
Solve_RationalBis q
Exit Sub
Else
Solve_Rational q
If Result_p.Count = 0 And Result_q.Count = 0 Then
        
        Msg.Add "this Method failed.You May try other Methods.Choose Method from List Below :"
        
        Msg.Add "(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu"
        flag = 1
       
        End If
Exit Sub
End If


'Exit Sub
End If
End If





If N = 0 Then
Msg.Add "Found a constant Polynomials !" + Chr(13) + Chr(10)
Exit Sub

End If



End Sub

Public Sub P_Message()
' *** Invoking Message Errors***
If P_Error = 1 Then
Msg.Add "Capacity Overflow !" + Chr(13) + Chr(10)
End If
If P_Error = 2 Then
Msg.Add "Attempt to divide by 0 !" + Chr(13) + Chr(10)
End If
If P_Error = 3 Then
Msg.Add "Division by Increasing power Impossible !" + Chr(13) + Chr(10)
End If
P_Error = 0

End Sub

Public Function P_Write(P As Collection) As String
'***Displaying A Polynomials***
Dim i As Long
Dim BPOLY As Boolean
Dim DG As Long
Dim Str As String
Dim tmp() As Variant
If P.Count = 0 Then
Str = "0"
Exit Function
End If

ReDim tmp(1 To P.Count)
For i = 1 To P.Count
tmp(i) = P(i)
Next

DG = Deg_Poly(P)

BPOLY = False
Str = ""
If IsNullPolynom(P) = True Then
Str = "0"
'Exit Function

Else
For i = DG + 1 To 1 Step -1
If tmp(i) <> 0 Then
    If BPOLY = True And (tmp(i) <> 0) Then
    If tmp(i) > 0 Then
    Str = Str + "+"
    Else
    Str = Str + "-"
    End If
    tmp(i) = Abs(tmp(i))
    
    
    End If
    BPOLY = True
    If (i = 1) Or (1 < i) And (Abs(tmp(i)) <> 1) Then
    Str = Str + " " + CStr(tmp(i))
    Else
    If tmp(i) = -1 Then
    Str = Str + "-"
    End If
    If i > 1 And (Abs(tmp(i)) <> 1) Then Str = Str + UCase(unkown)
    End If
    If i >= 2 Then
    Str = Str + UCase(" " & unkown) + " ^ " + CStr(i - 1)
    End If
    Str = Str + "  "
    'End If
    'End If
    End If
    
'End If
Next

End If
P_Write = Str

End Function

Public Sub Rational_Bis(P As Collection)
Dim N As Long
Dim i As Long
Dim j As Long
Dim Str As String
N = Deg_Poly(P)
Set R_p = New Collection
Set R_q = New Collection
Factor P(1), F_p
Factor P(N + 1), F_q
If (F_p.Count = 0 Or F_q.Count = 0) Then
Msg.Add "No factors found !" + Chr(13) + Chr(10)
Exit Sub

End If
     For i = 1 To F_p.Count
     For j = 1 To F_q.Count
     R_p.Add F_p(i)
     R_q.Add F_q(j)
     
     
     Next
     Next
    Resolve_Fac R_p
    Resolve_Fac R_q
    
    
  If R_p.Count <> 0 And R_q.Count <> 0 Then
Msg.Add "Investigating Possible Rational Roots.Second Method..."
Msg.Add "If p/q is a rational root for P(" + UCase(unkown) + ")" + Chr(13) + Chr(10)
Msg.Add "Possible solutions are: " + Chr(13) + Chr(10)

Str = "p = " & "{ "
For i = 1 To R_p.Count
Str = Str + CStr(R_p(i)) + " , "
Next
Str = Left(Str, Len(Str) - 3)
Str = Str + "  }" + Chr(13) + Chr(10)
Str = Str + Chr(13) + Chr(10) + "q = " & "{ "
For i = 1 To R_q.Count
Str = Str + CStr(R_q(i)) + " , "
Next
Str = Left(Str, Len(Str) - 3)
Str = Str + "  }" + Chr(13) + Chr(10)
Msg.Add Str
Exit Sub
End If
If R_p.Count = 0 And R_q.Count = 0 Then
Msg.Add "No Possible Rational Roots Found ." + Chr(13) + Chr(10)
Exit Sub

End If
End Sub

Public Sub P_Quadratic(P As Collection)
Msg.Add "Found a Quadratic Polynomial "
Msg.Add String$(60, "-")
On Error GoTo aa
'Get the data from the Parsed polynomial
Dim sf As String
Dim sg As String
Dim part1 As Double
Dim Desc As Double
Dim part2 As Double
Dim A, B, C As Double

A = Val(P(3))
B = Val(P(2))
C = Val(P(1))


part1 = -B / (2 * A)
Desc = B ^ 2 - 4 * A * C
Msg.Add "Descriminent = " + CStr(Desc)

'If the descriminant is negative then strip off the minus sign before calculating the square root
If Desc < 0 Then
Msg.Add "Negative Descriminent"
Msg.Add "Found Complex roots"
part2 = (-Desc) ^ 0.5 / (2 * A)
Else
If Desc > 0 Then
Msg.Add "Positive Descriminent"
Msg.Add "Found Real roots"
part2 = (Desc) ^ 0.5 / (2 * A)
End If
End If

If Desc = 0 Then
' results for real roots
Msg.Add "Null Descriminent"
Msg.Add "Found Unique Real Root"
Msg.Add "Root : " & UCase(unkown) + " = " + Format(part1, "#0.0####") + Chr(13) + Chr(10)
Msg.Add "Factorised Polynomial: "
Msg.Add "( " & UCase(unkown) + " - " + Format(part1, "#0.0####") & ")"
End If
If Desc > 0 Then
If (part1 + part2) <= 0 Then sg = "+"
If (part1 - part2) <= 0 Then sf = "+"
' results for real roots
Msg.Add "Root1 : " & UCase(unkown) + " = " + Format(part1 + part2, "#0.0####") + Chr(13) + Chr(10) + "Root2 : " & UCase(unkown) + " = " + Format(part1 - part2, "#0.0####") + Chr(13) + Chr(10)
Msg.Add "Factorised Polynomial: "
Msg.Add "( " & UCase(unkown) & Format(-part1 - part2, sg + "#0.0####") & ")*( " & UCase(unkown) & Format(-part1 + part2, sf + "#0.0####") & ") = 0"
End If
If Desc < 0 Then
If (part1) <= 0 Then sg = "+"
' results for imaginery roots roots
Msg.Add "Root1 : " & UCase(unkown) & " = " & Format(part1, "#0.0####") & " +j" & Format(part2, "#0.0####") + Chr(13) + Chr(10) + "Root2 : " & UCase(unkown) & " = " & Format(part1, "#0.0####") & " -j" & Format(part2, "#0.0####") + Chr(13) + Chr(10)
Msg.Add "Factorised Polynomial: "
Msg.Add "( " & UCase(unkown) & Format(-part1, sg + "#0.0####") & " -j" & Format(part2, "#0.0####") & " )( " & UCase(unkown) & Format(-part1, sg + "#0.0####") & " +j" & Format(part2, "#0.####") & ") = 0"
End If
Exit Sub

'In case of errors bail out of routine and clear windows
aa:
Msg.Add "Error :" + Err.Description
Exit Sub

End Sub

Public Sub P_Linear(P As Collection)
Msg.Add "Found a Linear Polynomial"
Msg.Add String$(60, "-")
Msg.Add "Equation :" + P_Write(P) + " =  0"
If P(1) = 0 Then
Msg.Add " Zero is the only possible root."
Exit Sub
Else
Msg.Add UCase(unkown) + " = - (" + CStr(P(1)) + "/" + CStr(P(Deg_Poly(P) + 1)) + ")"
Msg.Add "Unique Real root : " + UCase(unkown) + " = " + CStr(-Val(P(1)) / Val(P(Deg_Poly(P) + 1)))

End If

End Sub

Public Function R_VAL0(X As Variant) As Double
If Abs(X) <= 0 Then
R_VAL0 = 0
Else
R_VAL0 = X
End If

End Function

Public Sub P_Euclid(A As Collection, B As Collection, C As Collection, D As Collection)
'***Euclidien division Of Polynomials A and B *****
'***Quotient in C , Remainder In D    *************
Dim i As Long
Dim j As Long
Dim DA As Long
Dim DB As Long
Dim DC As Long
Dim PA() As Variant
Dim PB() As Variant
Dim PC() As Variant
Dim pd() As Variant
Set C = New Collection
Set D = New Collection

ReDim PA(1 To A.Count)
ReDim PB(1 To B.Count)
ReDim pd(1 To A.Count)

For i = 1 To A.Count
PA(i) = A(i)
Next
For i = 1 To B.Count
PB(i) = B(i)
Next
'Copy A in D
For i = 1 To A.Count
pd(i) = PA(i)
Next

    DA = Deg_Poly(A) + 1
    DB = Deg_Poly(B) + 1
    
    
    If DB < 0 Then
    P_Error = 2
    P_Message
    
    Exit Sub
    End If
    
    DC = DA - DB + 1
    If DC <= 0 Then
    Exit Sub
    End If
    ReDim PC(1 To DC)
     For i = 1 To UBound(PC)
     PC(i) = 0
     Next
    Do
    PC(DC) = pd(DC + DB - 1) / PB(DB)
    pd(DA) = 0
    If PC(DC) <> 0 Then
    For i = DC To DA - 1 Step -1
    
    pd(i) = pd(i) - PC(DC) * PB(i - DC + 1)
    Next
    End If
    DA = DA - 1
    DC = DC - 1
    
    Loop While DC >= 1
    
    
    ' End Euclid Fill C and D
    For i = 1 To UBound(PC)
    C.Add PC(i)
    Next
    For i = 1 To UBound(pd)
    
    D.Add pd(i)
   
    Next
    
End Sub

Public Sub SyntethicDivision(X As Collection, Y As Collection, C As Collection, D As Collection)


'Force declaration of variables


'routine to perform synthetic division
Dim q() As Double
Dim r() As Double
    Dim t() As Variant 'temporary array used so that the original X() doesn't get changed
    
    Dim oX As Long       'Order of the numerator
    Dim oY As Long       'Order of the denominator
    Dim j As Long       'used as a counter
    Dim i As Long        'used as a counter
    Dim ii As Long
    Dim jj As Long
    
    oX = Deg_Poly(X) + 1        'Order of X() is the size of the upper bound of the numerator array
    oY = Deg_Poly(Y) + 1        'Order of Y() is the size of the upper bound of the denominator array
    Set C = New Collection
    Set D = New Collection
    ReDim q(oX - oY + 1)
    ReDim r(oX)
    ReDim t(oX)
    For i = 1 To oX
        t(i) = X(i)
    Next i
    
    For i = 1 To oX - oY + 1
     
        q(i) = t(i) / Y(1)
        For j = 1 To oY
        
            t(j + i - 1) = t(j + i - 1) - q(i) * Y(j)
        Next j
    Next i
    
    For i = 1 To UBound(t)
        r(i) = t(i)
    Next i
    
    For i = 1 To UBound(q)
        C.Add q(i)
    Next i
    
    For i = 1 To UBound(q)
        D.Add r(i)
    Next i


End Sub

Public Function IsNullPolynom(P As Collection) As Boolean
Dim i As Long
Dim Bool As Boolean
Bool = True

For i = 1 To P.Count
If P(i) <> 0 Then
  Bool = False
  Exit For

End If
Next
IsNullPolynom = Bool

End Function

Public Sub Wait(Secs As Single)
On Error Resume Next

 Dim sngStart As Single

 sngStart = Timer()

 Do
 DoEvents
 
 Loop Until Timer() >= sngStart + Secs
 frmConsole.SetFocus
End Sub

Public Sub Solve_Rational(P As Collection)
'On Error GoTo ErrHandler
Dim i As Long
Dim j As Long
Dim factorised As String
Dim factorised2 As String
Dim D As Double



Dim ptmp As Collection
Dim tmp As Collection
Dim q As Collection
Dim r As Collection
Set q = New Collection
Set r = New Collection
Set ptmp = New Collection
Set Result_p = New Collection
Set Result_q = New Collection
Set tmp = New Collection

factorised = "Factorised Polynomial : "
For i = 1 To P.Count
ptmp.Add P(i)
Next

Rational_root P
If R_p.Count = 0 Or R_q.Count = 0 Then
Msg.Add "Rational Root First Method Failed"
Solve_RationalBis P

Exit Sub

End If

        'testing the Rationals
    For i = 1 To R_p.Count
    Set tmp = New Collection
    D = CDbl(R_p(i) / R_q(i))
    Msg.Add "Trying " + CStr(R_p(i)) + "/" + CStr(R_q(i)) + "  ..." + Chr(13) + Chr(10)
    
    If P_Value(P, D) <> 0 Then
    Msg.Add "p(" + CStr(R_p(i)) + "/" + CStr(R_q(i)) + ") = " + CStr(P_Value(P, D)) + "   Is Not Zero !" + Chr(13) + Chr(10)
    
    Msg.Add "Discarding " + CStr(R_p(i)) + "/" + CStr(R_q(i)) + Chr(13) + Chr(10)
    'End If
    
    Else
    tmp.Add Val(-R_p(i)) ' Dividor Polynom
    tmp.Add Val(R_q(i))
    ' try syntethic division of P and tmp
    P_Euclid ptmp, tmp, q, r
    Msg.Add CStr(P_Write(ptmp)) + " = (" + P_Write(tmp) + ")*(" + P_Write(q) + ") + (" + P_Write(r) + ")" + Chr(13) + Chr(10)
    
    If IsNullPolynom(r) = False Then
    Msg.Add CStr(Val(R_p(i))) + "/" + CStr(Val(R_q(i))) + "Remainder Not Zero !" + Chr(13) + Chr(10) + "Discarding " + CStr(R_p(i)) + "/" + CStr(R_q(i)) + Chr(13) + Chr(10)
    
    Else
    'If IsNullPolynom(R) = True Then ' found a rational Root
    Result_p.Add R_p(i) 'add the roots to the results
    Result_q.Add R_q(i)
    Msg.Add "Remainder is Zero -> Successful root : " + CStr(R_p(i)) + "/" + CStr(R_q(i)) + Chr(13) + Chr(10)
    
    Set ptmp = New Collection
    For j = 1 To q.Count  ' set the quotient as the P
    ptmp.Add q(j)
     Next
    
    End If
    End If
    Next
    If Result_p.Count = 0 And Result_q.Count = 0 Then
    Msg.Add "No Rational Roots found !" + Chr(13) + Chr(10)
    
    
    Exit Sub
    End If
    If Result_p.Count > 0 And Result_q.Count > 0 Then
    Msg.Add "Rational Roots : " + Chr(13) + Chr(10)
    For i = 1 To Result_p.Count
    Msg.Add "       " + "Root " + CStr(i) + " = " + CStr(Val(Result_p(i))) + "/" + CStr(Val(Result_q(i))) + Chr(13) + Chr(10)
    
    Next
    isRat = True
    
    'factorised = factorised + "(" + CStr(P_Write(Q)) + " )"
   ' msg = msg + "Factorised polynomial : "
  
  
  For i = 1 To Result_p.Count
  Set tmp = New Collection
  tmp.Add Val(-Result_p(i))
  tmp.Add Val(Result_q(i))
  factorised = factorised + " (" + P_Write(tmp) + ") * "
  factorised2 = factorised2 + " (" + P_Write(tmp) + ") * "
  Next
  
  factorised = factorised + "(" + P_Write(q) + ")" + Chr(13) + Chr(10)
  Msg.Add factorised
  Msg.Add "Finding roots for : " + P_Write(q)
  myIteration q
  If dRoots.Count > 0 Then
  For i = 1 To dRoots.Count
  factorised2 = factorised2 + "(" + UCase(unkown) + " - " + CStr(dRoots(i)) + ") * "
  Next
  factorised2 = Left(factorised2, Len(factorised2) - 3)
  Msg.Add String$(60, "-")
  Msg.Add "Factorised Polynomial Final Result :" + Chr(13) + Chr(10) + factorised2
  
  End If
    Exit Sub
    
    End If
'ErrHandler:
   ' MsgBox Err.Description, vbCritical
End Sub

Public Sub Initialize()
P_Error = 0
flag = 0
isRat = False
Set Msg = New Collection
Set Result_p = New Collection
Set Result_q = New Collection

Set F_p = New Collection
Set F_q = New Collection

Set F_pPlusq = New Collection
Set F_pMinusq = New Collection
Set F_pByq = New Collection
Set R_p = New Collection
Set R_q = New Collection
Set R_pByq = New Collection

End Sub

Public Sub Solve_RationalBis(P As Collection)
'On Error GoTo ErrHandler
Dim i As Long
Dim j As Long
Dim k As Long

Dim factorised As String
Dim factorised2 As String

Dim D As Double



Dim ptmp As Collection
Dim tmp As Collection
Dim q As Collection
Dim r As Collection
Set q = New Collection
Set r = New Collection
Set ptmp = New Collection
Set Result_p = New Collection
Set Result_q = New Collection
Set tmp = New Collection

factorised = "Factorised Polynomial : "
For i = 1 To P.Count
ptmp.Add P(i)
Next
Rational_Bis P


        'testing the Rationals
    For i = 1 To R_p.Count
    For k = 1 To R_q.Count
    Set tmp = New Collection
    D = CDbl(R_p(i) / R_q(k))
    Msg.Add "Trying " + CStr(R_p(i)) + "/" + CStr(R_q(k)) + "  ..." + Chr(13) + Chr(10)
    
    If P_Value(P, D) <> 0 Then
    Msg.Add "p(" + CStr(R_p(i)) + "/" + CStr(R_q(k)) + ") = " + CStr(P_Value(P, D)) + "   Is Not Zero !" + Chr(13) + Chr(10)
    
    Msg.Add "Discarding " + CStr(R_p(i)) + "/" + CStr(R_q(k)) + Chr(13) + Chr(10)
    'End If
    
    Else
    tmp.Add Val(-R_p(i)) ' Dividor Polynom
    tmp.Add Val(R_q(k))
    ' try syntethic division of P and tmp
    P_Euclid ptmp, tmp, q, r
    Msg.Add CStr(P_Write(ptmp)) + " = (" + P_Write(tmp) + ")*(" + P_Write(q) + ") + (" + P_Write(r) + ")" + Chr(13) + Chr(10)
    
    If IsNullPolynom(r) = False Then
    Msg.Add CStr(Val(R_p(i))) + "/" + CStr(Val(R_q(k))) + "Remainder Not Zero !" + Chr(13) + Chr(10) + "Discarding " + CStr(R_p(i)) + "/" + CStr(R_q(k)) + Chr(13) + Chr(10)
    
    Else
    'If IsNullPolynom(R) = True Then ' found a rational Root
    Result_p.Add R_p(i) 'add the roots to the results
    Result_q.Add R_q(k)
    Msg.Add "Remainder is Zero -> Successful root : " + CStr(R_p(i)) + "/" + CStr(R_q(k)) + Chr(13) + Chr(10)
    
    Set ptmp = New Collection
    For j = 1 To q.Count  ' set the quotient as the P
    ptmp.Add q(j)
     Next
    
    End If
    End If
    Next
    Next
     'Resolve_Fac Result_p
     'Resolve_Fac Result_q
    If Result_p.Count = 0 And Result_q.Count = 0 Then
    Msg.Add "No Rational Roots found !" + Chr(13) + Chr(10)
    
    Exit Sub
    End If
    If Result_p.Count > 0 And Result_q.Count > 0 Then
    Resolve_Fac Result_p
    Resolve_Fac Result_q
    Resolve_R Result_p, Result_q
    Msg.Add "Rational Roots : " + Chr(13) + Chr(10)
    For i = 1 To Result_p.Count
    Msg.Add "       " + "Root " + CStr(i) + " = " + CStr(Val(Result_p(i))) + "/" + CStr(Val(Result_q(i))) + Chr(13) + Chr(10)
    Next
    Set q = New Collection
    Set r = New Collection
    Set ptmp = New Collection
    For i = 1 To P.Count
    ptmp.Add P(i)
    Next
    For i = 1 To Result_p.Count
  Set tmp = New Collection
  tmp.Add Val(-Result_p(i))
  tmp.Add Val(Result_q(i))
  factorised = factorised + " (" + P_Write(tmp) + ") * "
  factorised2 = factorised2 + " (" + P_Write(tmp) + ") * "
  SyntethicDivision ptmp, tmp, q, r
  Set ptmp = New Collection
    For j = 1 To q.Count  ' set the quotient as the P
    ptmp.Add q(j)
     Next
  Next
  
  factorised = factorised + "(" + P_Write(q) + ")" + Chr(13) + Chr(10)
    Msg.Add factorised
  Msg.Add "Finding roots for : " + P_Write(q)
  myIteration q
  If dRoots.Count > 0 Then
  For i = 1 To dRoots.Count
  factorised2 = factorised2 + "(" + UCase(unkown) + " - " + CStr(dRoots(i)) + ") * "
  Next
  factorised2 = Left(factorised2, Len(factorised2) - 3)
  Msg.Add String$(60, "-")
  Msg.Add "Factorised Polynomial Final Result :" + Chr(13) + Chr(10) + factorised2
  
  End If
  
 Exit Sub
    End If
'ErrHandler:
    'MsgBox Err.Description, vbCritical
End Sub

Public Sub P_Bernoulli(P As Collection)
Dim N As Single
Dim i As Single
Dim PI As Double
Dim st As String
Dim s(MaxDeg) As Double, A(MaxDeg) As Double, C() As Single, t() As Single
Dim L As Double, Real As Double, imag As Double, r As Double, Z As Double, B As Double, aa As Double, bb As Double, cc As Double
Set Msg = New Collection
'msg.Add "Bernoulli Method of Moments..."

20 '
30 '
40  Msg.Add String$(60, "-")
50  'msg.Add "*                                       *"
60 'Print "* (c) 1986, Science Specialists Inc.   *"
70 'Print "*                                      *"
80  Msg.Add "        Polynomial Solving             "
90  Msg.Add String$(60, "-")
100 Msg.Add "    Bernoulli Method of Moment        "
101 'msg.Add "*                                       *"
110 'msg.Add String$(42, "-")

120  '*******************Input variables***************
130 N = Deg_Poly(P)

If N < 1 Then Exit Sub
140 '
150 For i = 0 To N
160 ' Print
165 A(i) = P(N - i + 1)
170 Next i
  '*******************display Equation***************
st = String$(60, "-"): Msg.Add st: st = ""
    For i = N To 0 Step -1:
    If A(N - i) >= 0 Then st = st + "+" + CStr(Abs(A(N - i))) + UCase(unkown) + "^" + CStr(i) + " " Else st = st + "-" + CStr(Abs(A(N - i))) + UCase(unkown) + "^" + CStr(i) + " "


 Next i

  st = st + "=0 ": Msg.Add st
  Msg.Add String$(60, "-")

' *******************Print Equation***************
171
'******************* Normalize***************
172 For i = 0 To N
175 If i > 0 Then A(i) = A(i) / A(0)
190 Next i: A(0) = 1
192 'FOR A=0 TO N:PRINT A,A(A):NEXT
'******************* Calculate Largest Moment ***************
195 s(0) = N - 1
200 For PI = 1 To 40: B = 0
210 If PI > N Then
L = N: 'Print "s4";:
GoTo 270
Else: L = PI - 1
End If
220 B = -PI * A(PI): 'Print pi, A(pi), b:
If PI = 1 Then GoTo 295
270 For Z = 1 To L
280 B = (-A(Z) * s(PI - Z)) + B
'Print b
290 Next Z
295 s(PI) = B: If Abs(s(PI)) > 1E+20 Then Exit For
400 Next PI
'******************* Print Real Roots ***************
410 'PRINT N
430 If (Sgn(s(PI - 1)) = Sgn(s(PI - 3))) And (Sgn(s(PI - 2)) = Sgn(s(PI - 4))) Then r = s(PI - 1) / s(PI - 2): Msg.Add "REAL ROOT=" + CStr(r): GoTo 1000: GoTo 200 Else GoTo 2000

1000  '******************* Reduce Equation ***************
1010 For i = 1 To N
1020 A(i) = A(i) + A(i - 1) * r
1025 Next
1027 If Abs(A(N)) > 0.001 Then Msg.Add "Warning: this method may not be suitable"
1030 Msg.Add "AFTER DIVISION-REMAINDER=" + CStr(A(N))
1035 ' msg.Add
1040 N = N - 1: If N = 1 Then Msg.Add "REAL ROOT=" + CStr(-A(1) / A(0)): GoTo 2222
1045 If N < 1 Then Exit Sub
1060 GoTo 172
' *******************Extract Complex Roots***************
2000 aa = (s(PI - 3) ^ 2 - s(PI - 2) * s(PI - 4))
2005 If aa = 0 Then Msg.Add "NO SOLUTION BY THIS METHOD": Exit Sub
2010 B = -s(PI - 2) * s(PI - 3) + s(PI - 1) * s(PI - 4): B = B / aa
2020 cc = ((s(PI - 2)) ^ 2 - s(PI - 3) * s(PI - 1)): cc = cc / aa: aa = 1
2025 Real = -B / (2 * aa): imag = Sqr(Abs(B ^ 2 - 4 * aa * cc)) / (2 * aa)
2030 Msg.Add "IMAGINARY ROOT=" + CStr(-B / (2 * aa)) + "   +J" + CStr(Sqr(Abs(B ^ 2 - 4 * aa * cc)) / (2 * aa))
2040 Msg.Add "IMAGINARY ROOT=" + CStr(-B / (2 * aa)) + "   -J" + CStr(Sqr(Abs(B ^ 2 - 4 * aa * cc)) / (2 * aa))
2050 aa = -2 * Real: B = Real ^ 2 + imag ^ 2
'******************* Reduce Equation ***************
2060 For i = 1 To N - 1
2070 A(i) = A(i) - A(i - 1) * aa
2080 A(i + 1) = A(i + 1) - A(i - 1) * B
2090 Next i
2091 If ((Abs(A(N)) > 0.001) Or (Abs(A(N - 1)) > 0.001)) Then Msg.Add "Warning: this method may not be suitable"
2092 If A(N) > 0 Then GoTo 2095 Else GoTo 2096
2095 Msg.Add "REMAINDER=(" + CStr(A(N - 1)) + "*" + UCase(unkown) + "+" + CStr(A(N)) + ")": GoTo 2097
2096 Msg.Add "REMAINDER=(" + CStr(A(N - 1)) + "*" + UCase(unkown) + " - " + CStr(Abs(A(N))) + ")"
2097 '
2100 N = N - 2
2104 If N = 1 Then Msg.Add "REAL ROOT=" + CStr(-A(1) / A(0)): GoTo 2222
2105 If N < 1 Then GoTo 2222
2110 GoTo 172
2115 If N = 0 Then GoTo 2222
2222 'PRINT"========================================":fa$=INPUTbox("NEW DATA"):IF INSTR("NoNOnOno",FA$)<1 THEN 20


End Sub

Public Sub P_Newton(P As Collection)
Dim flg As Integer, st As String

Dim er, aaaa$, rdm, aa, bb, ew, Real, imag, ei, wr1, wr2, D, dd
Dim PI, cr, ci, soq, iter, start, q, ang, sumr, rmdr, sumi, j, sr, Sq, N, soi, sor, flag
Dim A(20) As Single, B(20) As Single, s(20) As Single, RF As Single, jazz, cplex, rmd, r
Set Msg = New Collection



RF = 1
10 GoSub 1000
20 PI = 3.141592653589
30 GoSub 1120
35 'LOCATE 3, 52:
'Print "ITERATION"
40 '************** NEWTON REDUCTION ***************
70 sr = 0.2: cr = 1: ci = 0: flag = 0: start = 1: ' LOCATE 6, 52:
'Print "N"
80 GoSub 170: GoSub 270: soq = Sq
90 cr = cr + sr: GoSub 170: GoSub 270: iter = iter + 1
100 If soq < Sq Then flag = 0 Else flag = flag + 1
110 If flag = 0 Then If start = 1 Then start = 0 Else GoTo 400
140 sr = Sq / ((soq - Sq) / sr): If Abs(Sq - soq) < 0.000000000001 Then GoTo 260
145 '************** CALCULATE Y  *******************
150 If Abs(sr) < 0.000001 Then GoTo 260
160 soq = Sq: GoTo 90
170 sumr = 0: sumi = 0
180 For j = 0 To N
190 r = Sqr(cr ^ 2 + ci ^ 2)
200 If cr = 0 Then If ci > 0 Then ang = PI / 2 Else ang = -PI / 2 Else ang = Atn(ci / cr): If cr < 0 Then ang = ang + PI
210 sumr = A(j) * r ^ j * Cos(ang * j) + sumr
220 sumi = A(j) * r ^ j * Sin(ang * j) + sumi
230 Next j
240 Sq = Sqr(sumi ^ 2 + sumr ^ 2): 'LOCATE 4, 50:
'Text2.Text = "": ' LOCATE 4, 50:
'Text2.Text = "Iteration#: " & Format$(iter, "##########  ") & "Evaluation: " & Format$(cr, "##########  ") & ", " & Format$(ci, "##########  ")
250 Return
260 If Abs(Sq) > 0.000001 Then GoTo 290 Else GoTo 390
270 Sq = Sqr(sumr ^ 2 + sumi ^ 2): If Sq < 0.000001 Then GoTo 260 Else Return
280 soq = Sqr(sor ^ 2 + soi ^ 2): Return
285 '************** SEARCH FOR COMPLEX ROOT ********
290 sr = 0.2: ci = 1: iter = 0: start = 1: 'LOCATE 6, 52:
'Print "C"
300 GoSub 170: GoSub 270
310 soq = Sq: iter = iter + 1
320 ci = ci + sr: GoSub 170: GoSub 270
330 If soq < Sq Then flag = 0 Else flag = flag + 1
340 If flag = 0 Then sr = -sr / 2.0010010001
350 If flag > 2 Then sr = sr * 3.1415926535
360 If (Abs(soq - Sq) < 0.00000000001) Then GoTo 390
370 If Abs(sr) < 0.00000001 Then GoTo 390
380 GoTo 310
390 ' LOCATE 6, 52:
 GoSub 1330: GoTo 40
395 '************** NON-NEWTONIAN SEARCH ***********
400 sr = 0.2: flag = 0: iter = 0: Sq = 0: soq = 0: 'LOCATE 6, 52:
'Print "R"
410 cr = cr + sr: GoSub 170: iter = iter + 1
420 If soq < Sq Then flag = 0 Else flag = flag + 1
430 If flag = 0 Then sr = -sr / 2.1010100001
440 If flag > 3 Then sr = sr * 3.1415926535
450 If (Abs(Sq - soq) < 0.00000000001) Then GoTo 500
470 If Abs(sr) < 0.000000001 Then GoTo 260
480 soq = Sq: GoTo 410
490 sumr = 0: sumi = 0
500 GoSub 1260: If Abs(rmdr) > 0.000001 Then GoTo 290 Else GoSub 1320: GoTo 70
1000 ' Key OFF
1010  'ForeColor = RGB(0, 255, 0)
'Cls: Form1.ForeColor = RGB(255, 255, 255)
1020 Msg.Add String$(60, "-")
1030 'msg.Add "*                                       *"
1040 ' Print "* (c) 1986, Science Specialists Inc.   *"
1050 ' Print "*                                      *"
1060 Msg.Add "       Polynomial Solving             "
1070 Msg.Add String$(60, "-")
1080 Msg.Add "         Newton method             "
1090 'msg.Add "*                                       *"
1100 Msg.Add String$(60, "-")
1110 'ForeColor = (RGB(255, 255, 255)): Return
1120 N = Deg_Poly(P)
If N < 1 Then Exit Sub
1140 For q = N To 0 Step -1
1150 ' Print ""
1160 A(q) = P(q + 1): Next
1162 'RF = CSRLIN:
st = String$(60, "-"): Msg.Add st: st = ""
    For dd = N To 0 Step -1:
    If A(dd) >= 0 Then st = st + "+" & Abs(A(dd)) & UCase(unkown) + "^" & dd & " " Else st = st + "-" & Abs(A(dd)) & UCase(unkown) + "^" & dd & " "


 Next

  st = st + "=0 ": Msg.Add st
  Msg.Add String$(60, "-")
  'print st
If N = 2 Then GoTo 1800
1165 Return
1170 '********** reduce to standard form *********
1180 For q = N - 1 To 0 Step -1
1190 If q > 0 Then A(q) = A(q) / A(N)
1200 Next q: A(N) = 1
1220 s(0) = N - 1
1230 '********** REMAINDER CHECK     *************
1240 '
1250 '
1260  For j = 0 To N: B(j) = A(j): Next: jazz = 1: GoSub 1360: jazz = 0
1270  For j = 0 To N: A(j) = B(j): Next
1280  Return
1290    '
1320  '****************** PRINT OUT ROUTINE ********************
1330  cplex = 0: 'IS THE ROOT COMPLEX
1340  If Abs(ci) > 0.00001 Then cplex = 1: GoTo 1344
1341  GoSub 1260
1343  If Abs(rmdr) > 0.0001 Then cplex = 1 Else cplex = 0
1344  'LOCATE RF, 1:
If cplex = 1 Then GoSub 1430 Else er = cr: Msg.Add "REAL ROOT=" & CStr(er): RF = RF + 1: GoSub 1360: N = N - 1: GoSub 1400
1345  ci = 0: cr = 0: If N = 2 Then GoTo 1800 Else Return
1346 '
1350 '*** Reduce Equation by division ************
1360 For q = N - 1 To 0 Step -1
1370 A(q) = A(q) + A(q + 1) * cr
1380 Next: rmdr = A(0): rdm = rmdr: If jazz = 0 Then Msg.Add "     REMAINDER=" & CStr(rdm)
If Abs(rdm) > 0.001 Then Msg.Add "Warning : This method may not be suitable"
1390 For q = 1 To N: A(q - 1) = A(q): Next: Return
1400 If N = 1 Then
'LOCATE RF, 1:
ew = -A(0) / A(1): Msg.Add "REAL ROOT = " & CStr(ew): RF = RF + 1: GoTo 1700
End If
1410 If N < 1 Then GoTo 1700
1420 Return
1430 '******** extract imaginary roots ***********
1480 Real = cr: imag = ci
1490 'LOCATE RF, 1
 er = Real: ei = imag: Msg.Add "COMPLEX ROOT = " & CStr(er) & "   +J " & CStr(Abs(ei))
1500 Msg.Add "COMPLEX ROOT=" & CStr(er) & "   -J " & CStr(Abs(ei)): RF = RF + 2
1510 aa = -2 * Real: bb = Real ^ 2 + imag ^ 2
1520 '******** reduce equation size by division **
1530 For q = N - 1 To 1 Step -1
1540 A(q) = A(q) - A(q + 1) * aa
1550 A(q - 1) = A(q - 1) - A(q + 1) * bb
1560 Next q: A(N) = 1: wr1 = A(1): wr2 = A(0): Msg.Add " RMDR=(" + UCase(unkown) + "*" & CStr(wr1) & " " & CStr(wr2) & ")"
If (wr1 ^ 2 + wr2 ^ 2) ^ 0.5 > 0.001 Then Msg.Add "This method may not be suitable"
1570 N = N - 2: For q = 0 To N: A(q) = A(q + 2): Next
1580 If N = 1 Then
'LOCATE RF, 1:
er = A(0): ei = A(1): Msg.Add "REAL ROOT = " & CStr(-er / ei): GoTo 1700
End If
1590 If N < 1 Then GoTo 1700
1600 Return
1610 If N = 0 Then GoTo 1700
1700 'ForeColor = (RGB(0, 255, 0)): 'LOCATE  1,52:
Exit Sub: ' aaaa$ = InputBox("NEW DATA"): ForeColor = (RGB(255, 255, 255)): If InStr("nNnoNO", aaaa$) = 0 Then GoTo 10
1710 'LOCATE 1, 1:
End
1800 'LOCATE RF, 1
 D = A(1) ^ 2 - 4 * A(2) * A(0)
1900 If D >= 0 Then er = (-A(1) + Sqr(D)) / (2 * A(2)): ei = 0: GoSub 2000: er = (-A(1) - Sqr(D)) / (2 * A(2)): GoSub 2000
1910 If D < 0 Then er = (-A(1)) / (2 * A(2)): ei = Sqr(Abs(D)) / (2 * A(2)): GoSub 2100
1920 GoTo 1700
2000 Msg.Add "REAL ROOT=" & CStr(er): Return
2100 Msg.Add "COMPLEX ROOT = " & CStr(er) & "   +J " & CStr(Abs(ei))
2200 Msg.Add "COMPLEX ROOT = " & CStr(er) & "   -J " & CStr(Abs(ei)): Return


End Sub

Public Sub Bairstow(P As Collection)
 
 Dim i As Long
 Dim INITR As Boolean
 
 Dim PA(200), PB(200), PC(200) As Double
 Dim tPA As Collection
 Dim tPC As Collection
 Dim tmp As Collection
 Dim prod As Collection
 Dim r As Collection
 Dim SI, PI, S0, P0, C0, C1, C2, s, L, D As Double
 Dim k, N, CPTG, CPTI As Integer
 Dim Erreur As Boolean
 Set tPA = New Collection
 Set tPC = New Collection
 Set Msg = New Collection
 Set prod = New Collection
 Set tmp = New Collection
  Msg.Add String$(60, "-")
  Msg.Add "        Polynomial Solving             "
  Msg.Add String$(60, "-")
  Msg.Add "         Bairstow Method         "
  Msg.Add "  Irreductible Polynomial Factors  "
  Msg.Add String$(60, "-")
 

 
 INITR = False
 
 For i = 0 To Deg_Poly(P)
 PA(i) = P(i + 1)
 Next
 For i = 0 To Deg_Poly(P)
 PB(i) = PA(i)
 Next
 For i = 0 To Deg_Poly(P)
 PC(i) = 0
 Next
 N = Deg_Poly(P)
 CPTG = 0
 Erreur = False
 
        Do While (N > 1) And (CPTG <= MAXG)
        If INITR = True Then
        SI = -2# * PA(N - 1) / PA(N) / N
        If SI = 0# Then SI = 2#
        PI = 2# * PA(N - 2) / PA(N) / N / (N - 1)
        If PI = 0# Then PI = 1#
        
        Else
        SI = 2#
        PI = 1#
        CPTG = 0 'counter to zero
        CPTI = 0
        s = SI
        L = PI
         Do
         If Erreur Or (MAXI > 50) Then 'in case of divergence
         SI = SI + 0.2 * (Rnd - 0.5)
         PI = PI + 0.2 * (Rnd - 0.5)
         CPTG = CPTG + 1
         CPTI = 0
         s = SI
         L = PI
         
         End If
         If (L = 0#) And (s = 0#) Then L = 0.5 + Rnd
         S0 = s
         P0 = L
         CPTI = CPTI + 1
         PB(N) = PA(N)
         PB(N - 1) = PA(N - 1) + S0 * PB(N)
         C1 = 0
         C2 = PB(N)
         k = N - 2
         Do
         C0 = C1
         C1 = C2
         Erreur = DEBORDE(S0) Or DEBORDE(C1) Or DEBORDE(P0) Or DEBORDE(C0)
         If Not Erreur Then C2 = PB(k + 1) + S0 * C1 - P0 * C0
         Erreur = Erreur Or DEBORDE(C2)
         If Not Erreur Then PB(k) = PA(k) + S0 * PB(k + 1) - P0 * PB(k + 2)
         Erreur = Erreur Or DEBORDE(PB(k))
         k = k - 1
         
         Loop While Not Erreur And (k >= 0)
         If Not Erreur Then ' calculating S and L
         D = C2 * C0 - C1 * C1
         s = PB(0) * C0 - PB(1) * C1
         L = PB(0) * C1 - PB(1) * C2
         If D <> 0# Then
         s = S0 - s / D
         L = P0 - L / D
         End If
         D = Abs(S0) + Abs(P0)
         C0 = Abs(s - S0) + Abs(L - P0)
         If D > 1 Then C0 = C0 / D
         
         End If
         Loop While Erreur Or (MINR <= C0) And (MAXG >= CPTG)
         If CPTG <= MAXG Then
         D = R_VAL0(s * s - 4# * L) 'discriminent
         If D < 0# Then
         
         Set tPC = New Collection
         PC(2) = 1#
         PC(1) = -s
         PC(0) = L
         For i = 0 To 2
         tPC.Add PC(i)
         Next
        Msg.Add " ( " + P_Write(tPC) + " ) "
        
         Else
         
         Set tPC = New Collection
         PC(1) = 1#
         PC(0) = R_VAL0(-0.5 * (s + D ^ 0.5))
         For i = 0 To 1
         tPC.Add PC(i)
         Next
         Msg.Add " ( " + P_Write(tPC) + " ) "
          
         
         
         End If
         N = N - 2
         For k = 0 To N
         PA(k) = PB(k + 2)
         Next
         End If
        End If
        Loop
        If N > 0 Then
       
        Set tPA = New Collection
        
        For i = 0 To N
        tPA.Add PA(i)
        Next
        P_Normalise tPA
        Msg.Add " ( " + P_Write(tPA) + " ) "
        
        
        End If
        Msg.Add String$(60, "-")
       
        
        
End Sub

Public Function DEBORDE(ByVal X As Double) As Boolean
DEBORDE = (MAXR < Abs(X))

End Function

Public Sub P_Normalise(A As Collection)
Dim i, DA As Integer
Dim tA(MaxDeg) As Double
DA = Deg_Poly(A)
For i = 0 To DA
tA(i) = A(i + 1)

Next
 If DA >= 0 Then
 For i = 0 To DA - 1
 tA(i) = R_VAL0(tA(i) / tA(DA))
 Next
 tA(DA) = 1
 
 End If
 Set A = New Collection
 For i = 0 To DA
 A.Add tA(i)
 Next
 
End Sub

Public Sub P_Mul(A As Collection, B As Collection, C As Collection)
Dim i, j, k, DC As Long
Dim tA(MaxDeg) As Double
Dim tB(MaxDeg) As Double
Dim tC(MaxDeg) As Double

Set C = New Collection
For i = 0 To Deg_Poly(A)
tA(i) = A(i + 1)
Next
For i = 0 To Deg_Poly(B)
tB(i) = B(i + 1)
Next

For i = 0 To MaxDeg
tC(i) = 0
Next
DC = Deg_Poly(A) + Deg_Poly(B)
If DC > MaxDeg Then
P_Error = 1
P_Message
Exit Sub
End If
For k = 0 To DC
    For i = 0 To k
    tC(k) = tC(k) + tA(i) * tB(k - i)
    Next
Next
For i = 0 To DC
C.Add tC(i)
Next
End Sub

Public Function myIteration(ByVal P As Collection) As Integer
'Force declaration of all variables


   Dim Real() As Double       'Real part of this root
   Dim Cplx() As Double      'Complex part of this root

    
     Dim A() As Double
     Dim O As Long

    'Declare working variables
    Dim Iteration As Integer
    Dim N As Integer
    Dim Nx As Integer
    Dim Ny As Integer
    Dim RootCounter As Integer
    Dim L As Integer
    Dim i As Integer
    Dim in1 As Integer
    Dim ic As Integer

    Dim C() As Double
    Dim X As Double
    Dim Xo As Double
    Dim X2 As Double
    Dim Xp As Double
    Dim Xt As Double

    Dim Y As Double
    Dim Yo As Double
    Dim Y2 As Double
    Dim Yp As Double
    Dim Yt As Double
    
    Dim dX As Double
    Dim uX As Double
    Dim dY As Double
    Dim uY As Double
    Dim u As Double
    Dim v As Double
    Dim Sq As Double
    Dim AL As Double
    Dim t As Double

    'Dim Msg As String
    O = Deg_Poly(P)
    
    If isRat = False Then
    Set Msg = New Collection
    End If
    Set dRoots = New Collection
    
  Msg.Add String$(60, "-")
  Msg.Add "        Polynomial Solving             "
  Msg.Add String$(60, "-")
  Msg.Add "         Iteration Method         "
  Msg.Add String$(60, "-")
    ReDim A(0 To O)
    ReDim Real(0 To O - 1)
    ReDim Cplx(0 To O - 1)
    ReDim C(O)
    For i = 0 To O
    A(i) = P(i + 1)
    Next
        Iteration = 0
    RootCounter = 0

    N = O
    Nx = O
    Ny = O

'Reverse the order of the coefficients
    For L = 0 To O
        i = O - L
        C(i) = A(L)
    Next L

Rem: Set initial calculation values
lab15:  Xo = prfInitialGuess
        Yo = prfInitialGuess
        in1 = 0
        
lab8:   X = Xo
        Xo = -10 * Yo
        Yo = -10 * X
        X = Xo              'set X to current value
        Y = Yo              'set Y to current value
        in1 = in1 + 1
        GoTo lab3
        
lab10:  Iteration = 1
        Xp = X
        Yp = Y

Rem: Evaluate polynomial and derivatives
lab3:   ic = 0

lab7:   uX = 0
        uY = 0
        v = 0
        Yt = 0
        Xt = 1
        u = C(N)
        If u = 0 Then GoTo lab4

        For i = 1 To N
          L = N - i
          X2 = X * Xt - Y * Yt
          Y2 = X * Yt + Y * Xt
          u = u + C(L) * X2
          v = v + C(L) * Y2
          uX = uX + i * Xt * C(L)
          uY = uY - i * Yt * C(L)
          Xt = X2
          Yt = Y2
        Next i

        Sq = (uX ^ 2) + (uY ^ 2)
        If Sq = 0 Then GoTo lab5
        dX = (v * uY - u * uX) / Sq
        X = X + dX
        dY = -(u * uY + v * uX) / Sq
        Y = Y + dY
        If Abs(dY) + Abs(dX) - prfTolerance >= 0 Then
            ic = ic + 1
            If ic - prfMaxIterations < 0 Then GoTo lab7
            If Iteration = 0 Then
                If in1 - 5 < 0 Then
                    GoTo lab8
                Else
                    myIteration = ic
                    Exit Function
                End If
            End If
        End If

'Rem: Set the step iteration counter
lab6:   For L = 0 To Ny
          i = O - L
          t = A(i)
          A(i) = C(L)
          C(L) = t
        Next L
        i = N
        N = Nx
        Nx = i
        If Iteration <> 0 Then
            GoTo lab9
        Else
            GoTo lab10
        End If
        
lab5:   If Iteration = 0 Then GoTo lab8
        X = Xp
        Y = Yp

lab9:   Iteration = 0
        If Abs(Y) - (prfTolerance * Abs(X)) < 0 Then GoTo lab11
        AL = X + X
        Sq = (X ^ 2) + (Y ^ 2)
        N = N - 2
        GoTo lab12
        
lab4:   X = 0
        Nx = Nx - 1
        Ny = Ny - 1

lab11:  Y = 0
        Sq = 0
        AL = X
        N = N - 1
        
lab12:  C(1) = C(1) + AL * C(0)
        For L = 2 To N
          C(L) = C(L) + AL * C(L - 1) - Sq * C(L - 2)
        Next L


lab14:  Cplx(RootCounter) = Y
        Real(RootCounter) = X
        If Y = 0 Then
        Msg.Add "Real Root " + CStr(RootCounter + 1) + " = " + CStr(X)
        dRoots.Add CStr(X)
        
        Else
        If Y > 0 Then
        Msg.Add "Complex Root " + CStr(RootCounter + 1) + " = " + CStr(X) + " + J " + CStr(Abs(Y))
        dRoots.Add CStr(X) + " + J " + CStr(Abs(Y))
        Else
        Msg.Add "Complex Root " + CStr(RootCounter + 1) + " = " + CStr(X) + " - J " + CStr(Abs(Y))
        dRoots.Add CStr(X) + " - J " + CStr(Abs(Y))
        End If
        End If
        RootCounter = RootCounter + 1
        
        If Sq = 0 Then
            If N > 0 Then
                GoTo lab15
            Else
                myIteration = 0
                Exit Function
            End If
        Else
            Y = -Y
            Sq = 0
            GoTo lab14
        End If


End Function
