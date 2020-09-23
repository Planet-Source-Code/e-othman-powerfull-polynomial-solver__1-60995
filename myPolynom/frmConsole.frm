VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConsole 
   Caption         =   "Polynomial Solver - Console Mode"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cmndlg 
      Left            =   1080
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picContainer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      Begin VB.TextBox txtEquation 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS SystemEx"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuBc 
         Caption         =   "Backcolor"
      End
      Begin VB.Menu mnuFc 
         Caption         =   "ForColor"
      End
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EM_GETSEL = &HB0
Private Const EM_SETSEL = &HB1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEFROMCHAR = &HC9

Public oEqu As clsPolynom
 Dim P() As Variant

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
  
Private Sub Form_Load()

    Set oEqu = New clsPolynom
    picContainer.BackColor = txtEquation.BackColor
    Me.Show
    Me.Refresh
    txtEquation.SetFocus
    Reset txtEquation
    
    
End Sub

Private Sub Form_Resize()

    picContainer.Move 0, 0, ScaleWidth, ScaleHeight
    txtEquation.Move 60, 60, Abs(ScaleWidth - 50), Abs(ScaleHeight - 50)

End Sub

Private Sub Form_Unload(Cancel As Integer)

     Set oEqu = Nothing
     
End Sub



Private Sub mnuBc_Click()
On Error GoTo cancelling
Cmndlg.ShowColor

txtEquation.BackColor = Cmndlg.Color
picContainer.BackColor = txtEquation.BackColor
    Me.Show
    Me.Refresh
    txtEquation.SetFocus
cancelling:
Exit Sub
End Sub

Private Sub mnuFc_Click()
On Error GoTo cancelling
Cmndlg.ShowColor

txtEquation.ForeColor = Cmndlg.Color
picContainer.BackColor = txtEquation.BackColor
    Me.Show
    Me.Refresh
    txtEquation.SetFocus
cancelling:
Exit Sub
End Sub

Private Sub mnuQuit_Click()
Unload Me

End Sub

Private Sub txtEquation_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    If oEqu.flag = 0 Then
        Solve txtEquation
        Exit Sub
        End If
        If oEqu.flag = 1 Then
        Solve2 txtEquation
        Exit Sub
        End If
    End If
    
End Sub

Private Sub Solve(txt As TextBox)
'On Error GoTo Err_Expr
Dim i As Long

    'Dim P As Collection
    'Dim C As Collection
    'Dim D As Collection
    Dim Str As String
    Dim r() As Double
    Dim cursorPos           As Long
    Dim currLine            As Long
    Dim lineStartPos        As Long
    Dim sEquation           As String
    Dim lPos                As Long
    Dim sVar                As String
    Dim state1               As Boolean
    Dim state2               As Boolean
    Dim checkmsg            As String
    
    
    'On Error GoTo Err_Expr
   
    ' get the cursor position in the textbox
    cursorPos = SendMessage(txt.hwnd, EM_GETSEL, 0, ByVal 0&) \ &H10000
   
    ' get the current line index
    currLine = SendMessage(txt.hwnd, EM_LINEFROMCHAR, cursorPos, ByVal 0&)                           ' + 1
   
    ' get start of current line
    lineStartPos = SendMessage(txt.hwnd, EM_LINEINDEX, currLine, ByVal 0&)

    ' select the current lines text
    Call SendMessage(txt.hwnd, EM_SETSEL, lineStartPos, ByVal cursorPos)

    ' get the line
   
    
     sEquation = txt.SelText
     Str = sEquation
    ' reset selection to cursor position
    txt.SelStart = cursorPos
    
    If Len(sEquation) = 0 Then
    Exit Sub
    End If
    
    sEquation = Replace(sEquation, "?", vbNullString)
    
    lPos = InStr(sEquation, "=")
    If Left(LCase(sEquation), 4) = "exit" Or Left(LCase(sEquation), 4) = "quit" Then
    
    
    Unload Me
    Exit Sub
    
    End If
        ' assign a value to a variable
       ' sVar = Trim$(Mid$(sEquation, 1, lPos - 1))
       ' sEquation = Mid$(sEquation, lPos + 1)
        oEqu.Equation = sEquation
       ' oEqu.Solve
       ' oEqu.Var(sVar) = oEqu.Solution
        'End If
    state1 = IsPolynomial1(sEquation)
    oEqu.Parse_Poly
    state2 = IsPolynomial2(sEquation, oEqu.unkown)
   If state1 = False And state2 = False Then
    checkmsg = "Invalid Polynomial ! " + Chr(13) + Chr(10) + "Correct the syntax and try Again !"
    
    txt.SelLength = txt.SelStart
    txt.SelText = vbCrLf & CStr(checkmsg)
    Set oEqu.Coeff = New Collection
    Set oEqu.Pow = New Collection
    
    Else
    checkmsg = "Status : Valid Polynomial  "
    Set oEqu.Coeff = New Collection
    Set oEqu.Pow = New Collection
       
    txt.SelLength = txt.SelStart
    txt.SelText = vbCrLf & CStr(checkmsg)
    oEqu.Equation = sEquation
        
        
        
        
        oEqu.P_Read
        
        'oEqu.Var(oEqu.unkown) = 1
        
        'txt.SelLength = txt.SelStart
        'txt.SelText = vbCrLf & CStr(oEqu.msg)
        ' testing display
        'oEqu.P_Write oEqu.P
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("Parsed Polynomial : " & CStr(oEqu.P_Write(oEqu.P)))
        'txt.SelLength = txt.SelStart
        'txt.SelText = vbCrLf & CStr("Launching... ")
       
         
        ' oEqu.myIteration oEqu.P
         
        oEqu.Strategy_Launch
        
        For i = 1 To oEqu.Msg.Count
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr(oEqu.Msg(i))
        oEqu.Wait 1 / 5
        
        
        
        Next
        'Me.SetFocus
       
        
       
     ' calculate p(1)
     
       ' oEqu.Var(oEqu.unkown) = 1
         
       ' txt.SelLength = txt.SelStart
       ' txt.SelText = vbCrLf & CStr("P( 1)= " & oEqu.P_Value(oEqu.P, 1))
       ' txt.SelText = vbCrLf & CStr("P(-1)= " & oEqu.P_Value(oEqu.P, -1))
        
    
    'If lPos > 1 Then
        ' assign a value to a variable
       ' sVar = Trim$(Mid$(sEquation, 1, lPos - 1))
      '  sEquation = Mid$(sEquation, lPos + 1)
      '  oEqu.Equation = sEquation
      '  oEqu.Solve
      '  oEqu.Var(sVar) = oEqu.Solution
    'Else
        ' solve an equation
       ' oEqu.Equation = sEquation
        
        
        
       ' oEqu.Solve
       ' txt.SelLength = txt.SelStart
       ' txt.SelText = vbCrLf & CStr(oEqu.Solution)
    'End If
    End If
    If oEqu.flag = 0 Then
    Reset txt
    End If
    Exit Sub
    
'Err_Expr:
 'If Err.Number = 5020 Then
 ' checkmsg = "Invalid Polynomial ! " + Chr(13) + Chr(10) + "Correct the syntax and try Again !"
   ' txt.SelLength = txt.SelStart
  '  txt.SelText = vbCrLf & CStr(checkmsg)
   ' Set oEqu.Coeff = New Collection
  '  Set oEqu.Pow = New Collection
   ' Else
 
   ' MsgBox Err.Description
'End If
End Sub




Private Sub Solve2(txt As TextBox)
Dim i As Long

Dim cursorPos           As Long
    Dim currLine            As Long
    Dim lineStartPos        As Long
    Dim sEquation           As String
    Dim lPos                As Long
    Dim sVar                As String
    Dim state1               As Boolean
    Dim state2               As Boolean
    Dim checkmsg            As String
    Dim result As Integer
         'On Error GoTo Err_Expr
   
    ' get the cursor position in the textbox
    cursorPos = SendMessage(txt.hwnd, EM_GETSEL, 0, ByVal 0&) \ &H10000
   
    ' get the current line index
    currLine = SendMessage(txt.hwnd, EM_LINEFROMCHAR, cursorPos, ByVal 0&)                           ' + 1
   
    ' get start of current line
    lineStartPos = SendMessage(txt.hwnd, EM_LINEINDEX, currLine, ByVal 0&)

    ' select the current lines text
    Call SendMessage(txt.hwnd, EM_SETSEL, lineStartPos, ByVal cursorPos)

    ' get the line
         sEquation = txt.SelText
    
    ' reset selection to cursor position
    txt.SelStart = cursorPos
    
    If Len(sEquation) = 0 Then
    Exit Sub
    End If
    
    sEquation = Replace(sEquation, "?", vbNullString)
    
    lPos = InStr(sEquation, "=")
    If Left(LCase(sEquation), 4) = "exit" Or Left(LCase(sEquation), 4) = "quit" Then
    
    
    Unload Me
    Exit Sub
    
    End If
    Select Case Left(LCase(sEquation), 1)
        
        Case "i"
        result = oEqu.myIteration(oEqu.P)
        If result = 0 Then                  'if there were no errors
            For i = 1 To oEqu.Msg.Count
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr(oEqu.Msg(i))
        oEqu.Wait 1 / 5
        Next
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu")
        oEqu.flag = 1
        Me.SetFocus      'show the results display form
        Else
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("Unable to solve this equation in" + Str(result) + " Iterations." + Chr(13) + Chr(10) + "(N)ewton - (B)ernoulli - Bai(r)stow - (E)xit this Menu") 'otherwise  display and error
        oEqu.flag = 1
        Me.SetFocus
            
        End If
        
        Case "n"
        'oEqu.Equation = str
        
               
        
        'oEqu.P_Read
        oEqu.P_Newton oEqu.P
        For i = 1 To oEqu.Msg.Count
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr(oEqu.Msg(i))
        oEqu.Wait 1 / 5
        Next
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu")
        oEqu.flag = 1
        Me.SetFocus
        
        Case "r"
        'oEqu.Equation = str
        
               
        
        'oEqu.P_Read
        oEqu.Bairstow oEqu.P
        For i = 1 To oEqu.Msg.Count
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr(oEqu.Msg(i))
        oEqu.Wait 1 / 5
        Next
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration- (E)xit this Menu")
        oEqu.flag = 1
        Me.SetFocus
        
        Case "b"
        'oEqu.Equation = str
        
               
        
        'oEqu.P_Read
        oEqu.P_Bernoulli oEqu.P
        For i = 1 To oEqu.Msg.Count
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr(oEqu.Msg(i))
        oEqu.Wait 1 / 5
        Next
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("(N)ewton - (B)ernoulli - Bai(r)stow -(I)teration-(E)xit this Menu")
        oEqu.flag = 1
        Me.SetFocus
       
        Case "e"
        oEqu.flag = 0
        Reset txt
        
        Case Else
        txt.SelLength = txt.SelStart
        txt.SelText = vbCrLf & CStr("Invalid Choice !" + Chr(13) + Chr(10) + "(N)ewton - (B)ernoulli - Bai(r)stow - (E)xit this Menu")
        oEqu.flag = 1
        Me.SetFocus
        End Select
        'End If
End Sub

Public Sub Reset(txt As TextBox)
 Dim cursorPos           As Long
    Dim currLine            As Long
    Dim lineStartPos        As Long
    
    Dim lPos                As Long
    
     'On Error GoTo Err_Expr
   
    ' get the cursor position in the textbox
    cursorPos = SendMessage(txt.hwnd, EM_GETSEL, 0, ByVal 0&) \ &H10000
   
    ' get the current line index
    currLine = SendMessage(txt.hwnd, EM_LINEFROMCHAR, cursorPos, ByVal 0&)                           ' + 1
   
    ' get start of current line
    lineStartPos = SendMessage(txt.hwnd, EM_LINEINDEX, currLine, ByVal 0&)

    ' select the current lines text
    Call SendMessage(txt.hwnd, EM_SETSEL, lineStartPos, ByVal cursorPos)

    ' get the line
   
    
     
    
    ' reset selection to cursor position
    txt.SelStart = cursorPos
    txt.SelLength = txt.SelStart
    txt.SelText = vbCrLf & CStr("Type Polynomial Expression..." + Chr(13) + Chr(10))
    
End Sub
