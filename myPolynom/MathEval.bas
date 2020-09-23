Attribute VB_Name = "Module1"
'----------
'DISCLAIMER
'----------
'
'BY USING THIS SOURCE CODE, YOU INDICATE YOUR AGREEMENT TO THE FOLLOWING TERMS AND CONDITIONS:
'
'THE SOURCE CODE IS PROVIDED "AS IS" WITH NO REPRESENTATIONS OR WARRANTIES OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.  YOU ASSUME TOTAL RESPONSIBILITY AND RISK FOR YOUR USE OF THE SOURCE CODE.
'
'THE AUTHOR IS NOT RESPONSIBLE OR LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, SPECIAL, EXEMPLARY, PUNITIVE OR OTHER DAMAGES UNDER ANY CONTRACT, NEGLIGENCE, STRICT LIABILITY OR OTHER THEORY ARISING OUT OF OR RELATING IN ANY WAY TO THE SOURCE CODE. YOUR SOLE REMEDY FOR DISSATISFACTION WITH THE SOURCE CODE IS TO STOP USING THE SOURCE CODE.
'
'---------------------------------------
'JUST WHAT EXACTLY DOES THIS MODULE DO ?
'---------------------------------------
'
'One of the things I look for in each new version of Visual Basic is a function that will
'allow me to enter an algebraic formula as a string and will return the numeric value.
'Since this has not happened yet (and does not look like it will be part of VB 7),
'I spent a little time to create this module.
'
'The equation string to be solved can be any valid algebraic equation containing some or
'all of the following ASCII characters:
'
'      Valid operands are (in decreasing order of priority):
'          Parentheses "( )"
'          Exponent " ^ "
'          Division (with remainder) " / "
'          Division (drop remainder) " \ "
'          Multiplication " * "
'          Subtraction " - "
'          Addition " + "
'
'      Valid characters are:
'          Numerals "0-9"
'          Negative Sign             " - "
'          Thousands Separator       " , "
'          Decimal Separator         " . "
'
'---------------------------------------------
'CALLING THIS EXAMPLE CODE IN YOUR PROJECT ...
'---------------------------------------------
'
'    Dim dblSolution As Double
'
'    Debug.Print CStr(Calculate("(1+1)*(3+4)", dblSolution))
'    Debug.Print CStr(dblSolution)
'
'    Debug.Print ""
'
'    Debug.Print CStr(Calculate("1+1*3+4", dblSolution))
'    Debug.Print CStr(dblSolution)
'
'    Debug.Print ""
'
'    Debug.Print CStr(Calculate("(1+1)*3+4", dblSolution))
'    Debug.Print CStr(dblSolution)
'
'-----------------------------------------------------
'... WILL OUTPUT THE FOLLOWING TO THE IMMEDIATE WINDOW
'-----------------------------------------------------
'    True
'14
'
'    True
'8
'
'    True
'10
Public Function Calculate(ByVal Equation As String, ByRef Solution As Double) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
' OUTPUTS : Solution  --  A double-precision, floating-point number that contains
'                         the result of the algebraic formula provided.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Solution' will contain the result of the formula.
'           FAILURE   --  The function value will be FALSE.
'                         'Solution' will contain zero.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    '****   IF THE EQUATION IS VALIDATED, THEN CALCULATE A VALUE   ****
    If PrepareEquation(Equation) Then Calculate = FormattedCalc(Equation, Solution)

Exit Function

ErrorHandler:
    Calculate = False

End Function



Private Function FormattedCalc(ByRef Equation As String, ByRef Solution As Double) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
' OUTPUTS : Solution  --  A double-precision, floating-point number that contains
'                         the result of the algebraic formula provided.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Solution' will contain the result of the formula.
'           FAILURE   --  The function value will be FALSE.
'                         'Solution' will contain zero.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Dim lngOpen As Long
    Dim lngClose As Long
    Dim lngCount As Long
    Dim strChar As String
    Dim lngLoop As Long
    Dim dblResult As Double
    
    '*******************************************************************************
    '****   USE RECURSION TO SOLVE THE PARENTHETICAL PORTIONS OF THE EQUATION   ****
    '*******************************************************************************
    
    '****   INITIALIZE THE LOOP CONDITIONS   ****
    lngOpen = InStr(1, Equation, "(", vbTextCompare)
    Do While lngOpen > 0
        '****   INITIALIZE THE LOOP CONDITIONS   ****
        lngCount = 1
        lngLoop = lngOpen
        Do
            '****   INCRIMENT THROUGH THE EQUATION STRING   ****
            lngLoop = lngLoop + 1
            strChar = Mid$(Equation, lngLoop, 1)
            '****   IF WE FOUND NESTED PARENTHESES ...   ****
            If strChar = "(" Then
                '****   ... INCRIMENT THE PARENTHESES COUNTER ...   ****
                lngCount = lngCount + 1
            ElseIf strChar = ")" Then
                '****   .. OR DECRIMENT THE PARENTHESES COUNTER   ****
                lngCount = lngCount - 1
            End If
        '****   LOOP UNTIL WE FIND THE MATE OR RUN OUT OF STRING   ****
        Loop Until lngCount = 0 Or lngLoop = Len(Equation)
        
        '****   IF THE PARENTHESES MATCH UP ...   ****
        If lngCount = 0 Then
            '****   ... THEN MARK THE SPOT   ****
            lngClose = lngLoop
        Else
            '****   ... OTHERWISE WE HAVE AN UNMATCHED PARENTHESE   ****
            GoTo ErrorHandler
        End If
        
        '****   TREAT THE PARENTHETICAL STATEMENT AS AN ALL NEW SUB-EQUATION   ****
        If FormattedCalc(Mid$(Equation, lngOpen + 1, lngClose - lngOpen - 1), dblResult) Then
            '****   IF THE SUB-EQUATION WORKS, USE IT TO SIMPLIFY THIS EQUATION ...   ****
            Equation = Left$(Equation, lngOpen - 1) + CStr(dblResult) + _
                       Right$(Equation, Len(Equation) - lngClose)
        Else
            '****   OTHERWISE THIS EQUATION FAILS   ****
            GoTo ErrorHandler
        End If
        '****   LOOP UNTIL WE ARE ALL OUT OF PARENTHETICAL SUB-EQUATIONS   ****
        lngOpen = InStr(1, Equation, "(", vbTextCompare)
    Loop
    
    '****   SIMPLIFY THE BASIC OPERATORS IN ORDER OF PRECIDENCE   ****
    If Not SimplifyOperator(Equation, "^") Then GoTo ErrorHandler
    If Not SimplifyOperator(Equation, "/") Then GoTo ErrorHandler
    If Not SimplifyOperator(Equation, "\") Then GoTo ErrorHandler
    If Not SimplifyOperator(Equation, "*") Then GoTo ErrorHandler
    If Not SimplifyOperator(Equation, "-") Then GoTo ErrorHandler
    If Not SimplifyOperator(Equation, "+") Then GoTo ErrorHandler
        
    '****   RETURN THE SOLUTION   ****
    Solution = CDbl(Equation)
    FormattedCalc = True

Exit Function

ErrorHandler:
    '****   RETURN FAILURE   ****
    Solution = 0
    FormattedCalc = False

End Function



Private Function SimplifyOperator(ByRef Equation As String, ByVal Operator As String) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
'           Operator  --  A string character defining which operation to perform.
' OUTPUTS : Equation  --  The string parameter 'Equation' will be modified.  The
'                         simplified portions of the string will be replaced with
'                         the calculated equalities.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Equation' will contain the new, simplified equation.
'           FAILURE   --  The function value will be FALSE.
'                         'Equation' will be an empty string.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Dim lngOper As Long
    Dim intOper As Integer
    Dim dblOperOne As Double
    Dim dblOperTwo As Double
    Dim dblResult As Double
    Dim blnLeftOper As Boolean
    Dim blnRightOper As Boolean

    '****   INITIALIZE THE LOOP CONDITIONS   ****
    lngOper = InStr(1, Equation, Operator, vbTextCompare)
    intOper = Asc(Operator)
    Do While lngOper > 0
        '****   GET THE OPERANDS FOR THE SPECIFIED OPERATOR   ****
        blnLeftOper = GetOperand(Equation, lngOper, dblOperOne, True)
        blnRightOper = GetOperand(Equation, lngOper, dblOperTwo, False)
        '****   IF WE GET TWO VALID OPERATORS ...   ****
        If blnLeftOper And blnRightOper Then
            '****   ... THEN CALCULATE & SUBSTITUTE   ****
            Select Case intOper
                Case 94
                    dblResult = dblOperOne ^ dblOperTwo
                Case 47
                    dblResult = dblOperOne / dblOperTwo
                Case 92
                    dblResult = dblOperOne \ dblOperTwo
                Case 42
                    dblResult = dblOperOne * dblOperTwo
                Case 45
                    dblResult = dblOperOne - dblOperTwo
                Case 43
                    dblResult = dblOperOne + dblOperTwo
                Case Else
                    GoTo ErrorHandler
            End Select
            '****   IF THE SUBSTITUTION FAILS ...   ****
            If Not SubstituteAnswer(Equation, lngOper, dblResult) Then
                '****   ... THEN RETURN AN ERROR   ****
                GoTo ErrorHandler
            End If
            '****   LOOK FOR MORE OF THE SAME OPERATOR   ****
            lngOper = InStr(1, Equation, Operator, vbTextCompare)
            
        ElseIf Not blnLeftOper And blnRightOper And intOper = 45 Then
            '****   ... OTHERWISE IT MIGHT BE A NEGATIVE SIGN CHARACTER   ****
            lngOper = InStr(lngOper + 1, Equation, Operator, vbTextCompare)
            
        Else
            '****   ... IF NOTHING ELSE, RETURN AN ERROR   ****
            GoTo ErrorHandler
        End If
    Loop
    
    '****   RETURN SUCCESS   ****
    SimplifyOperator = True
    
Exit Function

ErrorHandler:
    '****   DESTROY THE EQUATION AND RETURN FAILURE   ****
    Equation = ""
    SimplifyOperator = False

End Function



Private Function SubstituteAnswer(ByRef Equation As String, _
                                  ByVal Location As Long, _
                                  ByVal Answer As Double) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
'           Location  --  A long integer that contains the location of the
'                         operator of the segemnt to be replaced.
'           Answer    --  A double-precision, floating-point variable containing
'                         the value with which to simplify the equation.
' OUTPUTS : Equation  --  The string parameter 'Equation' will be modified.  The
'                         simplified portion of the string identified by 'Location'
'                         will be replaced with the calculated equality 'Answer'.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Equation' will contain the new, simplified equation.
'           FAILURE   --  The function value will be FALSE.
'                         'Equation' will be an empty string.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Dim lngStart As Long
    Dim lngStop As Long
    Dim intChar As Integer
    Dim dblJunk As Double
    
    '****   INITIALIZE THE LOOP CONDITIONS   ****
    lngStart = Location - 1
    intChar = Asc(Mid$(Equation, lngStart, 1))
    
    '****   LOOP UNTIL WE FIND THE START OF THE OPERAND   ****
    Do While (intChar >= 48 And intChar <= 57) Or intChar = 44 Or intChar = 46
        '****   DECRIMENT THE LOOP COUNTER   ****
        lngStart = lngStart - 1
        '****   IF WE HIT THE END OF THE STRING ...   ****
        If lngStart < 1 Then
            '****   ... WE ARE DONE LOOPING   ****
            intChar = 0
        Else
            '****   ... OTHERWISE READ THE NEXT CHARACTER   ****
            intChar = Asc(Mid$(Equation, lngStart, 1))
        End If
    Loop
    '****   FIGURE OUT IF A LEADING "-" IS AN OPERAND OR A SIGN   ****
    If intChar = 45 And Not GetOperand(Equation, lngStart, dblJunk, True) Then
        '****   IF IT IS A SIGN, INCLUDE IT IN OUR OUTPUT VALUE   ****
        lngStart = lngStart - 1
    End If
    
    '****   INITIALIZE THE LOOP CONDITIONS   ****
    lngStop = Location + 1
    intChar = Asc(Mid$(Equation, lngStop, 1))
        
    '****   IF THE OPERATOR IS FOLLOWED BY A "-" SIGN ...   ****
    If intChar = 45 Then
        '****   ... SKIP OVER IT TO THE OPERAND   ****
        lngStop = lngStop + 1
        intChar = Asc(Mid$(Equation, lngStop, 1))
    End If
    
    '****   LOOP UNTIL WE FIND THE END OF THE OPERAND   ****
    Do While (intChar >= 48 And intChar <= 57) Or intChar = 44 Or intChar = 46
        '****   INCRIMENT THE LOOP COUNTER   ****
        lngStop = lngStop + 1
        '****   IF WE HIT THE END OF THE STRING ...   ****
        If lngStop > Len(Equation) Then
            '****   ... WE ARE DONE LOOPING   ****
            intChar = 0
        Else
            '****   ... OTHERWISE READ THE NEXT CHARACTER   ****
            intChar = Asc(Mid$(Equation, lngStop, 1))
        End If
    Loop
    
    '****   SIMPLIFY THE EQUATION AND RETURN SUCCESS   ****
    Equation = Left$(Equation, lngStart) + CStr(Answer) + _
               Right$(Equation, Len(Equation) - lngStop + 1)
    SubstituteAnswer = True
    
Exit Function

ErrorHandler:
    '****   DESTROY THE EQUATION AND RETURN FAILURE   ****
    Equation = ""
    SubstituteAnswer = False

End Function



Private Function GetOperand(ByVal Equation As String, _
                            ByVal Location As Long, _
                            ByRef Operand As Double, _
                            Optional ByVal LeftSide As Boolean = True) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
'           Location  --  A long integer that contains the location of the
'                         operator for which we want an operand.
'           LeftSide  --  A boolean that determins whether to return the operand
'                         from the left side or the right side of the operator.
' OUTPUTS : Operand   --  A double-precision, floating-point variable that contains
'                         numeric value of the requested operand.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Solution' will contain the requested operand.
'           FAILURE   --  The function value will be FALSE.
'                         'Solution' will contain zero.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Dim intChar As Integer
    Dim lngLoop As Long
    Dim intOffset As Integer
    Dim dblJunk As Double
    
    If LeftSide Then
        '****   LOOP TO THE LEFT   ****
        intOffset = -1
    Else
        '****   LOOP TO THE RIGHT   ****
        intOffset = 1
    End If
    
    '****   INITIALIZE THE LOOP CONDITIONS   ****
    lngLoop = Location + intOffset
    intChar = Asc(Mid$(Equation, lngLoop, 1))
    
    '****   IF THE OPERATOR IS FOLLOWED BY A "-" SIGN ...   ****
    If intChar = 45 Then
        '****   ... SKIP OVER IT TO THE OPERAND   ****
        lngLoop = lngLoop + 1
        intChar = Asc(Mid$(Equation, lngLoop, 1))
    End If
    
    '****   LOOP UNTIL WE FIND THE END OF THE OPERAND   ****
    Do While (intChar >= 48 And intChar <= 57) Or intChar = 44 Or intChar = 46
        '****   INCRIMENT/DECRIMENT THE LOOP COUNTER   ****
        lngLoop = lngLoop + intOffset
        '****   IF WE HIT THE END OF THE STRING ...   ****
        If lngLoop < 1 Or lngLoop > Len(Equation) Then
            '****   ... WE ARE DONE LOOPING   ****
            intChar = 0
        Else
            '****   ... OTHERWISE READ THE NEXT CHARACTER   ****
            intChar = Asc(Mid$(Equation, lngLoop, 1))
        End If
    Loop
    
    If LeftSide Then
        '****   FIGURE OUT IF A LEADING "-" IS AN OPERAND OR A SIGN   ****
        If intChar = 45 And Not GetOperand(Equation, lngLoop, dblJunk, True) Then
            '****   IF IT IS A SIGN, INCLUDE IT IN OUR OUTPUT VALUE   ****
            lngLoop = lngLoop - 1
        End If
        '****   RETURN THE OPERAND TO THE LEFT OF THE OPERATOR   ****
        Operand = CDbl(Mid$(Equation, lngLoop + 1, Location - lngLoop - 1))
        GetOperand = True
    Else
        '****   RETURN THE OPERAND TO THE RIGHT OF THE OPERATOR   ****
        Operand = CDbl(Mid$(Equation, Location + 1, lngLoop - Location - 1))
        GetOperand = True
    End If
    
Exit Function

ErrorHandler:
    '****   DESTROY THE OPERAND AND RETURN FAILURE   ****
    Operand = 0
    GetOperand = False
    
End Function




Private Function PrepareEquation(ByRef Equation As String) As Boolean
    
'--------------------------------------------------------------------------------
' INPUTS  : Equation  --  A string value that contains a valid algebraic formula
'                         as defined in the comments at the top of the module.
' OUTPUTS : Equation  --  The string parameter 'Equation' will be modified.  All
'                         whitespace will be removed.
' RETURN  : SUCCESS   --  The function value will be TRUE.
'                         'Equation' will contain the re-formatted equation.
'           FAILURE   --  The function value will be FALSE.
'                         'Equation' will be an empty string.
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    Dim lngLoop As Long
    Dim intChar As Integer
    Dim lngParen As Long
    
    '****   INITIALIZE THE LOOP CONDITIONS   ****
    Equation = Trim$(Equation)
    lngLoop = 1
    
    '****   LOOP THROUGH THE ENTIRE EQUATION   ****
    Do While lngLoop <= Len(Equation)
        '****   CHECK CHARACTERS BY ASCII VALUE   ****
        intChar = Asc(Mid$(Equation, lngLoop, 1))
        Select Case intChar
            Case 8, 9, 10, 13, 32
                '****   REMOVE ANY WHITESPACE   ****
                Equation = Mid$(Equation, 1, lngLoop - 1) + Mid$(Equation, lngLoop + 1)
            Case 94, 47, 92, 42, 45, 43, 48 To 57, 44 To 46
                '****   SKIP ANY VALID OPERATORS AND CHARACTERS   ****
                lngLoop = lngLoop + 1
            Case 40
                '****  INCRIMENT THE PARENTHESES COUNTER   ****
                lngParen = lngParen + 1
                '****   SKIP THE VALID OPERATOR   ****
                lngLoop = lngLoop + 1
            Case 41
                '****   DECRIMENT THE PARENTHESES COUNTER   ****
                lngParen = lngParen - 1
                '****   SKIP THE VALID OPERATOR   ****
                lngLoop = lngLoop + 1
            Case Else
                '****   ANYTHING ELSE INVALIDATES THE EQUATION   ****
                GoTo ErrorHandler
        End Select
        '****   YOU SHOULD NEVER HAVE MORE CLOSED   ****
        '****   PARENTHESES THAN OPEN PARENTHESES   ****
        If lngParen < 0 Then Exit Do
    Loop
    
    '****   IF EVERY OPEN PARENTHESE HAD A CLOSE PARENTHESE ...   ****
    If lngParen = 0 Then
        '****   ... RETURN SUCCESS   ****
        PrepareEquation = True
    Else
        '****   ... OTHERWISE WE HAVE UNMATCHED PARENTHESES   ****
        GoTo ErrorHandler
    End If
    
Exit Function

ErrorHandler:
    '****   DESTROY THE EQUATION AND RETURN FAILURE   ****
    Equation = ""
    PrepareEquation = False

End Function

