Attribute VB_Name = "Sudoku"
Option Explicit
Dim possibleValues() As String
Dim solutionLog() As String
Dim i As Integer 'x position
Dim j As Integer 'y position

Public Sub clearTab()
    ActiveSheet.Cells(11, 1).Resize(9, 9).Value = ""
End Sub

Public Sub SudokuSolver()
    Dim randomGuess As Boolean
    '''
    randomGuess = True
    '''
    ReDim solutionLog(0)
    ReDim possibleValues(1 To 9, 1 To 9)
    createNewPossibilitesArray
    If checkValuesFromTable = False Then
        Exit Sub
    Else
    End If

    For i = 1 To 9
        For j = 1 To 9
            If checker(i, j, False) = False Then
                If possibleValues(i, j) = "" And solutionLog(0) = "" Then
                    MsgBox ("hatalý puzzle")
                    Exit Sub
                Else
                    guesser i, j, randomGuess
                End If
            Else
            End If
        Next j
    Next i
    solution
End Sub

Private Sub guesser(i As Integer, j As Integer, randomGuess As Boolean)
    'randomly guess from possible solutions
    Dim aRndValue As String
    If randomGuess Then
        aRndValue = VBA.Mid(possibleValues(i, j), WorksheetFunction.Odd((Int(Rnd * (Len(possibleValues(i, j)) - 2)) + 1)), 1)
    Else
        aRndValue = VBA.Mid(possibleValues(i, j), 1, 1)
    End If
    Dim cmdString As String
    cmdString = "@" & i & "," & j & "V" & possibleValues(i, j) & "->" & aRndValue
    doCommand (cmdString)
End Sub

Private Function checker(i As Integer, j As Integer, startCheck As Boolean) As Boolean
    
    ActiveSheet.Cells(i, j).Interior.Color = 14348258
    'check it's viable numbers then return if detected or guessable value
    Dim xLineValues As String
    Dim yLineValues As String
    Dim squareValues As String
    If startCheck = True Then

        xLineValues = getValuesXLine(i)
        yLineValues = getValuesYLine(j)
        squareValues = getSquareVals(i, j)
        
        If Len(xLineValues & yLineValues & squareValues) - _
            Len(Replace(xLineValues & yLineValues & squareValues, possibleValues(i, j), "")) = 3 Then
            checker = True
        Else
            checker = False
        End If
    ElseIf Len(possibleValues(i, j)) = 1 Then
        checker = True
    Else
        xLineValues = getValuesXLine(i)
        yLineValues = getValuesYLine(j)
        squareValues = getSquareVals(i, j)
        
        If possibleValues(i, j) = "" Then possibleValues(i, j) = "1,2,3,4,5,6,7,8,9"
        
        'destroy unfitted values
        Dim a As Integer
        For a = 1 To Len(xLineValues) Step 2
            possibleValues(i, j) = Replace(possibleValues(i, j), VBA.Mid(xLineValues, a, 1), "")
        Next a
        For a = 1 To Len(yLineValues) Step 2
            possibleValues(i, j) = Replace(possibleValues(i, j), VBA.Mid(yLineValues, a, 1), "")
        Next a
        For a = 1 To Len(squareValues) Step 2
            possibleValues(i, j) = Replace(possibleValues(i, j), VBA.Mid(squareValues, a, 1), "")
        Next a

        'replace the ,,'s
        Do While InStr(possibleValues(i, j), ",,") <> 0
            possibleValues(i, j) = Replace(possibleValues(i, j), ",,", ",")
        Loop
        If VBA.Left(possibleValues(i, j), 1) = "," Then possibleValues(i, j) = VBA.Right(possibleValues(i, j), Len(possibleValues(i, j)) - 1)
        If VBA.Right(possibleValues(i, j), 1) = "," Then possibleValues(i, j) = VBA.Left(possibleValues(i, j), Len(possibleValues(i, j)) - 1)
        
        'Not possible solution left need to revert.
        If possibleValues(i, j) = "" Then
            If solutionLog(0) = "" Then
                            
            Else
                'Revert to allpossible state.
                possibleValues(i, j) = "1,2,3,4,5,6,7,8,9"
                'revert to one before
                revertLastCommand
            End If
        Else
        End If
        
        checker = False
        
    End If
    
    ActiveSheet.Cells(i, j).Interior.Color = 16777215

End Function

Private Sub revertLastCommand()
'   "@1,2V2,3,4,5->4"
'   "@1,2V2,3,4,5->4"
    Dim cmdString As String
    cmdString = solutionLog(UBound(solutionLog))

    Dim fromVal As String
    Dim toVal As String
    
    i = VBA.Mid(cmdString, 2, 1)
    j = VBA.Mid(cmdString, 4, 1)
    fromVal = VBA.Mid(cmdString, 6, InStr(cmdString, "->") - InStr(cmdString, "V") - 1)
    toVal = VBA.Mid(cmdString, InStr(cmdString, "->") + 2, Len(cmdString) - InStr(cmdString, "->"))
    'do command
    fromVal = Replace(fromVal, toVal, "")
    fromVal = Replace(fromVal, ",,", ",")
    If VBA.Left(fromVal, 1) = "," Then fromVal = VBA.Right(fromVal, Len(fromVal) - 1)
    If VBA.Right(fromVal, 1) = "," Then fromVal = VBA.Left(fromVal, Len(fromVal) - 1)
    
    possibleValues(i, j) = fromVal
    'log Command
    If UBound(solutionLog) = 0 Then
        If solutionLog(0) = "" And possibleValues(i, j) = "" Then Stop
        solutionLog(0) = ""
    Else
        ReDim Preserve solutionLog(0 To (UBound(solutionLog) - 1))
    End If
    'solutionLog(UBound(solutionLog)) = cmdString
    Debug.Print "---reverted---" & cmdString
    If fromVal = "" Then revertLastCommand
End Sub
Private Function getSquareVals(i As Integer, j As Integer)
    Dim sqRow As Integer
    Dim sqCol As Integer
    sqRow = (i + 2) \ 3
    sqCol = (j + 2) \ 3
    
    Dim a As Integer
    Dim b As Integer
    'If i = 7 And j = 7 Then Stop
    For a = (3 * sqRow - 2) To (sqRow * 3)
        For b = (3 * sqCol - 2) To (sqCol * 3)
            If ActiveSheet.Cells(a, b).Value <> "" Then
                getSquareVals = getSquareVals & "," & ActiveSheet.Cells(a, b).Value
            Else
            End If
            If InStr(possibleValues(a, b), ",") = 0 And InStr(getSquareVals, possibleValues(a, b)) = 0 And possibleValues(a, b) <> "" Then
                getSquareVals = getSquareVals & "," & possibleValues(a, b)
            Else
            End If
        Next b
    Next a
    'If Left(getSquareVals, 1) = ",," Then Stop
    If getSquareVals <> "" Then getSquareVals = VBA.Right(getSquareVals, Len(getSquareVals) - 1)
    'If Left(getSquareVals, 1) = "," Then Stop
End Function
Private Function getValuesXLine(i As Integer)
    Dim b As Integer
    For b = 1 To 9
        If ActiveSheet.Cells(i, b).Value <> "" Then
            getValuesXLine = getValuesXLine & "," & ActiveSheet.Cells(i, b).Value
        Else
        End If
        If InStr(possibleValues(i, b), ",") = 0 And InStr(getValuesXLine, possibleValues(i, b)) = 0 And possibleValues(i, b) <> "" Then
            getValuesXLine = getValuesXLine & "," & possibleValues(i, b)
        Else
        End If
    Next b
    If getValuesXLine <> "" Then getValuesXLine = VBA.Right(getValuesXLine, Len(getValuesXLine) - 1)
End Function

Private Function getValuesYLine(j As Integer)
    Dim a As Integer
    For a = 1 To 9
        If ActiveSheet.Cells(a, j).Value <> "" Then
            getValuesYLine = getValuesYLine & "," & ActiveSheet.Cells(a, j).Value
        Else
        End If
        If InStr(possibleValues(a, j), ",") = 0 And InStr(getValuesYLine, possibleValues(a, j)) = 0 And possibleValues(a, j) <> "" Then
            getValuesYLine = getValuesYLine & "," & possibleValues(a, j)
        Else
        End If
    Next a
    If getValuesYLine <> "" Then getValuesYLine = VBA.Right(getValuesYLine, Len(getValuesYLine) - 1)
End Function

Private Sub doCommand(cmdString As String)
'   "@1,2V2,3,4,5->4"
'   "@1,2V2,3,4,5->4"

    Dim i As Integer
    Dim j As Integer
    Dim fromVal As String
    Dim toVal As String
    
    i = VBA.Mid(cmdString, 2, 1)
    j = VBA.Mid(cmdString, 4, 1)
    fromVal = VBA.Mid(cmdString, 6, InStr(cmdString, "->") - InStr(cmdString, "V") - 1)
    toVal = VBA.Mid(cmdString, InStr(cmdString, "->") + 2, Len(cmdString) - InStr(cmdString, "->"))
    'do command
    If Len(fromVal) >= 2 And InStr(fromVal, ",") = 0 Then Stop
    possibleValues(i, j) = toVal
    'log Command
    If UBound(solutionLog) = 1 Or solutionLog(0) <> "" Then
        ReDim Preserve solutionLog(0 To (UBound(solutionLog) + 1))
        solutionLog(UBound(solutionLog)) = cmdString
    Else
        solutionLog(0) = cmdString
    End If
    Debug.Print cmdString

End Sub

Private Function checkValuesFromTable()
   
    For i = LBound(possibleValues, 1) To UBound(possibleValues, 1)
        For j = LBound(possibleValues, 2) To UBound(possibleValues, 2)
            If ActiveSheet.Cells(i, j).Value <> "" Then
                possibleValues(i, j) = CStr(ActiveSheet.Cells(i, j).Value)
                If checker(i, j, True) = False Then
                    MsgBox ("hatalý puzzle")
                    checkValuesFromTable = False
                    Exit Function
                Else
                    'possibleValues(i, j) = CStr(ActiveSheet.Cells(i, j).Value)
                End If
            Else
            End If
        Next j
    Next i
    checkValuesFromTable = True
End Function

Private Sub createNewPossibilitesArray()

    For i = LBound(possibleValues, 1) To UBound(possibleValues, 1)
        For j = LBound(possibleValues, 2) To UBound(possibleValues, 2)
            'If Len(possibleValues(i, j)) = 1 Then
                possibleValues(i, j) = "1,2,3,4,5,6,7,8,9"
            'Else
            'End If
        Next j
    Next i
End Sub

Private Sub solution()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 9
        For j = 1 To 9
            ActiveSheet.Cells(i + 10, j).Value = possibleValues(i, j)
            ActiveSheet.Cells(i + 10, j).Font.ColorIndex = possibleValues(i, j)
        Next j
    Next i
End Sub
