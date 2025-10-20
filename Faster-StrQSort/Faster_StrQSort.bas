Attribute VB_Name = "Faster_StrQSort"
' Copyright Ken J. Anton 2025

''# VBA Faster String Quick Sort
'' * This module provides a method for sorting a single dimension String array.
'' It uses a tweaked version of Quick Sort by replacing direct String swappings with indirect String pointer swappings.
'' It also utilizes some additinal speed optimizations.

''It works on either **`32/64-bit`**, and on any **`(MS Office`** versions and platforms (**`Windows`** or **`Mac`**).
''It also works on **`TwinBasic`**.

Option Explicit

' Declaring xMoveByRef for direct API call for all versions. It is intended for internal use only

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Sub xMoveByRef Lib "/usr/lib/libc.dylib" Alias "memmove" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
    #Else
        Private Declare  Sub xMoveByRef Lib "/usr/lib/libc.dylib" Alias "memmove" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As Long)
    #End If
#Else
    #If VBA7 Or TwinBasic Then
        Private Declare PtrSafe Sub xMoveByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
    #Else
        Private Declare  Sub xMoveByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As Long)
    #End If
#End If

#If VBA7 = 0 And TwinBasic = 0 Then
    ' This LongPtr type definition is for legacy 32-bit platform only.
    ' In case this type has been defined somewhere else in another module,
    '   this definition must be removed to prevent 'ambiguous type' error!
    
    Public Enum LongPtr ' By default, Enum type value is a Long type
        [_] ' Empty Enum
    End Enum
    
#End If

' Data ptr and SafeArray Bound (Packed together)
Public Type SA_DataRange
    p As LongPtr
    Cnt As Long
    IndexLB As Long
End Type

' Alternate SafeArray structure
Public Type SA_Type1
   Dims As Integer
   FeatureFlags As Integer
   eSize As Long
   LockCnt As Long
   DataRg As SA_DataRange
End Type

' SortOrder Enum for either Excel or non-Excel platforms
Public Enum QsSortOrder
    qsAscending = 1 ' same value as Excel's xlAscending
    qsDescending = 2 ' same value as Excel's xlDescending
End Enum

Public Sub StrQSort(Arr$(), _
            Optional ByVal SortOrder As QsSortOrder = qsAscending, _
            Optional ByVal CmpMethod As VbCompareMethod = vbBinaryCompare, _
            Optional ByVal Start& = &H80000000, Optional ByVal Last& = &H7FFFFFFF)
    
#If Win64 Then
    Const PtrLen& = 8
#Else
    Const PtrLen& = 4
#End If

    Const SA_DataPtr_Offset& = 8 + PtrLen, SA_DataRange_Length = 8 + PtrLen

    Dim pArr() As LongPtr, Arr_SAptr As LongPtr, pArr_SAptr As LongPtr
    Dim TmpSA_Hdr As SA_Type1
    
    If Last <= Start Then Exit Sub
    
    ReDim pArr(0 To 0)
    
' Check for empty array
    Arr_SAptr = Not Not Arr
    If Arr_SAptr = 0 Then Exit Sub
    
' Check array dims = 1
    xMoveByRef TmpSA_Hdr, ByVal Arr_SAptr, PtrLen ' it needs only to chk 'Dims' field
    If TmpSA_Hdr.Dims <> 1 Then
        Err.Raise 5, , "StrQSort does not support multi-dimension array"
        Exit Sub
    End If
        
' Validate Start and Last
    If Start < LBound(Arr) Then Start = LBound(Arr)
    If Last > UBound(Arr) Then Last = UBound(Arr)
    
' Save SafeArray of pArr
    pArr_SAptr = Not Not pArr
    xMoveByRef TmpSA_Hdr, ByVal pArr_SAptr, LenB(TmpSA_Hdr)
    
' Tweaking the pArr by copying DataPtr and Array Bounds of Arr to pArr
    xMoveByRef ByVal pArr_SAptr + SA_DataPtr_Offset, ByVal Arr_SAptr + SA_DataPtr_Offset, LenB(TmpSA_Hdr.DataRg)
    
' Do Sort
    zStrQSort Arr, pArr, Start, Last, (SortOrder = qsDescending) Or 1, CmpMethod ' CmpGT (1= Ascending, -1= Descending)
    
' Undo tweaking
    xMoveByRef ByVal pArr_SAptr, TmpSA_Hdr, LenB(TmpSA_Hdr)
End Sub

Private Sub zStrQSort(Arr$(), pArr() As LongPtr, ByVal Li&, ByVal Ri&, ByVal CmpGT&, ByVal CmpMethod As VbCompareMethod)
    Dim Pivot$, Lx&, Rx&, Mx&, R4&, Rh&, a_cmp_b&, a_cmp_c&, pTmp As LongPtr

    Rh = Ri - Li:  If Rh > 4 Then GoTo zzQSort
    
    Mx = Li + 1
    On Rh GoTo zz2, zz3, zz4, zz5
    Exit Sub
'--------
zzQSort:
    Lx = Li:  Rx = Ri:  Pivot = Arr((Li + Ri) \ 2)
    
    Do While Lx <= Rx
        Do While StrComp(Pivot, Arr(Lx), CmpMethod) = CmpGT:  Lx = Lx + 1:  Loop ' Lx++ until out of order
        Do While StrComp(Arr(Rx), Pivot, CmpMethod) = CmpGT:  Rx = Rx - 1:  Loop ' Rx-- until out of order
        
        If Lx > Rx Then Exit Do
        pTmp = pArr(Lx):  pArr(Lx) = pArr(Rx):  pArr(Rx) = pTmp ' swap Lx, Rx
        Lx = Lx + 1:  Rx = Rx - 1
    Loop
    
    zStrQSort Arr, pArr, Li, Rx, CmpGT, CmpMethod
    zStrQSort Arr, pArr, Lx, Ri, CmpGT, CmpMethod
    Exit Sub
'--------
zz5:
    Rx = Ri
    GoSub zzPreCmp3

    If StrComp(Arr(Rh), Arr(R4), CmpMethod) <> CmpGT Then ' [Rh <= R4]
        If StrComp(Arr(R4), Arr(Rx), CmpMethod) <> CmpGT Then GoTo zz3_sort    ' [Rh <= R4 <= Rx]
        pTmp = pArr(Rx):  pArr(Rx) = pArr(R4):  pArr(R4) = pTmp ' swap R4, Rx -> [Rh <= R4 and R4 > Rx]
        GoTo zz4_cmp_Rh_R4
    End If
    
    If StrComp(Arr(Rh), Arr(Rx), CmpMethod) <> CmpGT Then GoTo zz4_swap_Rh_R4 ' [Rh <= Rx]
    pTmp = pArr(Rx):  pArr(Rx) = pArr(Rh):  pArr(Rh) = pTmp ' swap Rh, Rx

zz4:
'--------
    GoSub zzPreCmp3
    
zz4_cmp_Rh_R4:
    If StrComp(Arr(Rh), Arr(R4), CmpMethod) <> CmpGT Then GoTo zz3_sort ' [Rh <= R4]
    
zz4_swap_Rh_R4:
    pTmp = pArr(R4):  pArr(R4) = pArr(Rh):  pArr(Rh) = pTmp ' swap Rh, R4

'--------
zz3:
    a_cmp_c = StrComp(Arr(Li), Arr(Ri), CmpMethod)
    a_cmp_b = StrComp(Arr(Li), Arr(Mx), CmpMethod)
    
zz3_sort:
    If a_cmp_c <> CmpGT Then ' if [a <= c]
        If a_cmp_b = CmpGT Then GoTo zzSwap_Li_Mx ' [a > b]  -> M K Z -> swap a, b
        If StrComp(Arr(Mx), Arr(Ri), CmpMethod) = CmpGT Then GoTo zzSwap_Mx_Ri ' [b > c]  -> K Z M -> swap  b, c
        Exit Sub ' [b <= c] -> K M Z -> in order
    Else ' [a > c]
        If a_cmp_b = -CmpGT Then GoTo zzRotate3_Right ' [b > a]  -> M Z K -> rotate right
        If StrComp(Arr(Ri), Arr(Mx), CmpMethod) = CmpGT Then GoTo zzRotate3_Left ' [c > b]  -> Z K M -> rotate left
        GoTo zzSwap_Li_Ri ' [b >= c]  -> Z M K -> swap a, c
    End If
'--------
zz2:  If StrComp(Arr(Li), Arr(Ri), CmpMethod) <> CmpGT Then Exit Sub ' [a <= b]

zzSwap_Li_Ri:  pTmp = pArr(Li):  pArr(Li) = pArr(Ri):  pArr(Ri) = pTmp:  Exit Sub
zzSwap_Li_Mx:  pTmp = pArr(Li):  pArr(Li) = pArr(Mx):  pArr(Mx) = pTmp:  Exit Sub
zzSwap_Mx_Ri:  pTmp = pArr(Mx):  pArr(Mx) = pArr(Ri):  pArr(Ri) = pTmp:  Exit Sub
'--------
zzRotate3_Left:  pTmp = pArr(Li):  pArr(Li) = pArr(Mx):  pArr(Mx) = pArr(Ri):  pArr(Ri) = pTmp:  Exit Sub
zzRotate3_Right:  pTmp = pArr(Ri):  pArr(Ri) = pArr(Mx):  pArr(Mx) = pArr(Li):  pArr(Li) = pTmp:  Exit Sub
'--------
zzPreCmp3:
    Ri = Mx + 1:  R4 = Ri + 1
            
    a_cmp_b = StrComp(Arr(Li), Arr(Mx), CmpMethod)
    a_cmp_c = StrComp(Arr(Li), Arr(Ri), CmpMethod)
    
    If a_cmp_b = CmpGT Then ' [a > b]
        If a_cmp_c <> CmpGT Then Rh = Ri Else Rh = Li ' if [a <= c]
    Else ' [a <= b]
        If StrComp(Arr(Mx), Arr(Ri), CmpMethod) = CmpGT Then Rh = Mx Else Rh = Ri ' if [b > c]
    End If
    Return
End Sub
