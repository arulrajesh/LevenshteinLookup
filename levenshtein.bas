Attribute VB_Name = "Module1"
Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

  Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
  Dim string1_length As Long
  Dim string2_length As Long
  Dim distance() As Long
  Dim min1 As Long, min2 As Long, min3 As Long

  string1_length = Len(string1)
  string2_length = Len(string2)
  ReDim distance(string1_length, string2_length)
  bs1 = string1
  bs2 = string2

  For i = 0 To string1_length
      distance(i, 0) = i
  Next

  For j = 0 To string2_length
      distance(0, j) = j
  Next

  For i = 1 To string1_length
      For j = 1 To string2_length
          'slow way: If Mid$(string1, i, 1) = Mid$(string2, j, 1) Then
          If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then   ' *2 because Unicode every 2nd byte is 0
              distance(i, j) = distance(i - 1, j - 1)
          Else
              'distance(i, j) = Application.WorksheetFunction.Min _
              (distance(i - 1, j) + 1, _
               distance(i, j - 1) + 1, _
               distance(i - 1, j - 1) + 1)
              ' spell it out, 50 times faster than worksheetfunction.min
              min1 = distance(i - 1, j) + 1
              min2 = distance(i, j - 1) + 1
              min3 = distance(i - 1, j - 1) + 1
              If min1 <= min2 And min1 <= min3 Then
                  distance(i, j) = min1
              ElseIf min2 <= min1 And min2 <= min3 Then
                  distance(i, j) = min2
              Else
                  distance(i, j) = min3
              End If

          End If
      Next
  Next

  Levenshtein = distance(string1_length, string2_length)

  End Function


Function LevenshteinVLookup(searchValue As String, searchRange As Range, colIndex As Integer, Optional maxDist As Integer = 2) As Variant
    Dim i As Long, minDist As Integer, curDist As Integer
    Dim matchRow As Long, matchFound As Boolean
    
    matchFound = False
    minDist = maxDist + 1
    
    For i = 1 To searchRange.Rows.Count
        curDist = Levenshtein(searchValue, searchRange.Cells(i, 1).Value)
        If curDist <= maxDist And curDist < minDist Then
            minDist = curDist
            matchRow = i
            matchFound = True
        End If
    Next i
    
    If matchFound Then
        LevenshteinVLookup = searchRange.Cells(matchRow, colIndex).Value
    Else
        LevenshteinVLookup = CVErr(xlErrNA)
    End If
End Function

 
