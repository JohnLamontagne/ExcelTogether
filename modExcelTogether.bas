Public Function COMPLEXSEPERATE(ByVal sepMask As String, ByVal source As String, ByVal index As Long) As String
    COMPLEXSEPERATE = ComplexSeperateString(sepMask, source, index, 0)
End Function

Private Function ComplexSeperateString(ByVal sepMask As String, ByVal source As String, ByVal rIndex As Long, ByVal cIndex As Long) As String
Dim LenText As Long, Pos As Long
Dim leftText As String
Dim rightText As String
    
    cIndex = cIndex + 1
    LenText = Len(source)
    Pos = 1
    
    Do
        s = Mid(source, Pos, Len(sepMask))
        If s Like sepMask Then
            
            
            leftText = Left(source, Pos - 1)
            rightText = Mid(source, Pos + Len(sepMask))
            
            
            
            If cIndex = rIndex Then
                ComplexSeperateString = leftText
                Exit Function
            Else
                ComplexSeperateString = ComplexSeperateString(sepMask, rightText, rIndex, cIndex)
                Exit Function
            End If

        End If
        Pos = Pos + 1

    Loop While Pos < (LenText - Len(sepMask) - 1)
    
    ' If there's nothing left to process in the string but the rIndex is higher than cIndex, something wen't wrong (either in seperate mask formatting or in users assumption
    ' of how many parts there would be
    If rIndex > cIndex Then
        ComplexSeperateString = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' If we get here, return whatever is left over (right-most part).
    ComplexSeperateString = source
    
End Function

