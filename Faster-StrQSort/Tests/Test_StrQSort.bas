Attribute VB_Name = "Test_StrQSort"
' Copyright Ken J. Anton 2025

Option Explicit

Public Sub Testing_StrQSort( _
            Optional ByVal SortOrder As QsSortOrder = qsAscending, _
            Optional ByVal CmpMethod As VbCompareMethod = vbBinaryCompare)
            
    Const SrcCnt& = 128
    Dim Arr$(), Line1$, Line2$, i&, n&, r&, CmpGT&
    ReDim Arr(1 To SrcCnt)
    
    Create_RandomText_Array Arr, SrcCnt
    
' Randomly assign duplicate lines
    For i = 1 To 8
        r = 1 + Int(Rnd * SrcCnt)
        Line1 = Arr(r)
        For n = 1 To 1 + Int(Rnd * 3)
            r = 1 + Int(Rnd * SrcCnt)
            Arr(r) = Line1
        Next
    Next
    
    StrQSort Arr, SortOrder, CmpMethod
    
    CmpGT = (SortOrder = qsDescending) Or 1 ' -> (1= Ascending, -1= Descending)

    Line1 = Arr(1)
    DebugPrint 1 & "- " & Line1
    
' Verify sorted result
    For i = 2 To SrcCnt
        Line2 = Arr(i)
        DebugPrint i & "- " & Line2
        Debug.Assert StrComp(Line1, Line2, CmpMethod) <> CmpGT ' Assert Line1 <= line2
        Line1 = Line2
    Next
    DebugPrint vbCrLf & "*** Testing successful !"
End Sub

Private Function RandomText$()
    Const Asc0 = 48, Asc9 = 57, AscA = 65, AscZ = 90, Asc_a = 97, Asc_z = 122
    Dim n&, i&, r&
    n = Int(Rnd * 64)
    RandomText = Space$(n)
    For i = 1 To n
        Do
            r = Asc0 + Int(Rnd * 75)
        Loop Until r <= Asc9 Or r >= AscA And r <= AscZ Or r >= Asc_a And r <= Asc_z
        If i = 1 And r <= Asc9 Then r = r + (AscA - Asc0)
        Mid$(RandomText, i, 1) = Chr$(r)
    Next
End Function

Private Sub Create_RandomText_Array(ByRef Arr$(), Optional ByVal Count&)
    Dim i&
    If Count <= 0 Then Count = 32
    ReDim Arr(1 To Count)
    For i = 1 To Count
        Arr(i) = RandomText
    Next
End Sub

Private Sub DebugPrint(Txt$)
    #if TwinBasic Then
        Console.Writeline Txt
    #else
        Debug.Print Txt
    #end if
End Sub