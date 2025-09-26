Attribute VB_Name = "DirectMem_Test"
Option Explicit

' The following Compiler Constants can be modified to allow groups of methods to be enabled(=1) or disabled(=0)

#Const Enable_MoveMem = 1

#Const Enable_GetMem = 1
#Const Enable_PutMem = 1
#Const Enable_PropertyGet = 1
#Const Enable_PropertyLet = 1

#Const Enable_Ptr = 1
#Const Enable_Byte = 1
#Const Enable_Int = 1
#Const Enable_Long = 1
#Const Enable_Cur = 1
#Const Enable_Sng = 1
#Const Enable_Dbl = 1
#Const Enable_Date = 1

#If Win64 Or TWINBASIC Then
    #Const Enable_LLong = 1
#End If

Private mDestAddr As LongPtr, mSrcAddr As LongPtr, mOverlapDestAddr As LongPtr, mOverlapSrcAddr As LongPtr

Public Sub Test_All_Except_MoveMem()
    Test_PtrAccess
    Test_ByteAccess
    Test_IntAccess
    Test_LongAccess
    Test_CurAccess
    Test_SngAccess
    Test_DblAccess
    Test_DateAccess
    Test_LLongAccess
End Sub

' Test MoveMem
' *****************

Public Sub Test_MoveMem()
    Const n1k = 1024&
    Const ub& = 256 * n1k, ExtraSpace& = 64 * n1k, HalfExtraSpace = ExtraSpace \ 2
    
    Dim SrcByteArr(0 To ub + 1024) As Byte, DestByteArr(0 To ub + ExtraSpace) As Byte, nBytes&, i&
    Dim vOffset, OverlapArr, StartByte&, EndByte&, StartOffset&, EndOffset&, LastOffset&
    
    #If Enable_MoveMem Then
                
        ' Fill Src arr with random byte values
        For i = 0 To ub
            SrcByteArr(i) = Rnd * 128
        Next i
        
        mSrcAddr = VarPtr(SrcByteArr(0))
        mDestAddr = VarPtr(DestByteArr(0))
        
        OverlapArr = Array(1, 4, 8, 16, 32, 64, 128, 256, 1024, 2048, 4096, 8192, 16 * n1k, 32 * n1k)
        
        Debug.Print "Testing MoveMem:  Number bytes being moved (from .. to ..) ** Gaps (DestAddr - SrcAddr) ..."; vbCrLf
        
        For StartByte = 0 To 96 * n1k Step 1024
            EndByte = StartByte + 1024 - 1
            
            StartOffset = IIf(StartByte <= 0, 1, 8)
            EndOffset = EndByte + 1
            
            Debug.Print "  from"; StartByte; "to"; EndByte; " ** Gaps:";
                        
            LastOffset = 0
            For Each vOffset In OverlapArr
                If vOffset > HalfExtraSpace Then vOffset = HalfExtraSpace
                
                If vOffset > EndOffset Then
                    If LastOffset = 0 Then
                        LastOffset = vOffset
                        EndOffset = vOffset
                    End If
                End If
                
                If (vOffset >= StartOffset And vOffset <= EndOffset) Then
                    Debug.Print vOffset; " ";
                    DoEvents
                    
                    For nBytes = StartByte To EndByte
                        Erase DestByteArr
                        
                        MoveMem mDestAddr, mSrcAddr, nBytes
                        
                        ' Asserting Src and Dest Arr are equal
                        For i = 0 To nBytes - 1
                            Debug.Assert SrcByteArr(i) = DestByteArr(i)
                        Next
            
                        ' Testing overlap moves
                        mOverlapSrcAddr = mDestAddr
                        mOverlapDestAddr = mDestAddr + vOffset
                    
                        MoveMem mOverlapDestAddr, mOverlapSrcAddr, nBytes
                    
                        ' Asserting Src and Dest Arr are equal
                        For i = 0 To nBytes - 1
                            Debug.Assert SrcByteArr(i) = DestByteArr(i + vOffset)
                        Next
                    Next
                End If
            Next
            
            Debug.Print "   ... completed"
        Next
        Debug.Print
        zSuccess "Test_MoveMem"
        
    #End If
End Sub

' Ptr (LongPtr) memory access
' *******************************

Public Sub Test_PtrAccess()
    Dim mDestPtr As LongPtr, mSrcPtr As LongPtr

    mDestAddr = VarPtr(mDestPtr)
    mSrcAddr = VarPtr(mSrcPtr)
    
    #If Enable_Ptr Then
        
        #If Enable_PropertyGet Then
            
            mDestPtr = 0
            mSrcPtr = 1234567890
            mDestPtr = PtrMem(mSrcAddr)
            zAssert mDestPtr, mSrcPtr, "mDestPtr = PtrMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestPtr = 0
            mSrcPtr = 1234567890
            PtrMem(mDestAddr) = mSrcPtr
            zAssert mDestPtr, mSrcPtr, "PtrMem(mDestAddr) = mSrcPtr"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestPtr = 0
            mSrcPtr = 1234567890
            GetMemPtr mSrcAddr, mDestPtr
            zAssert mDestPtr, mSrcPtr, "GetMemPtr mSrcAddr, mDestPtr"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestPtr = 0
            mSrcPtr = 1234567890
            PutMemPtr mDestAddr, mSrcPtr
            zAssert mDestPtr, mSrcPtr, "PutMemPtr mDestAddr, mSrcPtr"
    
        #End If
        
        zSuccess "Test_PtrAccess"
    
    #End If
        
End Sub
    
' Byte memory access
' **********************

Public Sub Test_ByteAccess()
    Dim mDestByte As Byte, mSrcByte As Byte

    mDestAddr = VarPtr(mDestByte)
    mSrcAddr = VarPtr(mSrcByte)
    
    #If Enable_Byte Then
        
        #If Enable_PropertyGet Then
            
            mDestByte = 0
            mSrcByte = 123
            mDestByte = ByteMem(mSrcAddr)
            zAssert mDestByte, mSrcByte, "mDestByte = ByteMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestByte = 0
            mSrcByte = 123
            ByteMem(mDestAddr) = mSrcByte
            zAssert mDestByte, mSrcByte, "ByteMem(mDestAddr) = mSrcByte"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestByte = 0
            mSrcByte = 123
            GetMem1 mSrcAddr, mDestByte
            zAssert mDestByte, mSrcByte, "GetMem1 mSrcAddr, mDestByte"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestByte = 0
            mSrcByte = 123
            PutMem1 mDestAddr, mSrcByte
            zAssert mDestByte, mSrcByte, "PutMem1 mDestAddr, mSrcByte"
    
        #End If
        
        zSuccess "Test_ByteAccess"
        
    #End If
        
End Sub

' Int memory access
' **********************

Public Sub Test_IntAccess()
    Dim mDestInt As Integer, mSrcInt As Integer

    mDestAddr = VarPtr(mDestInt)
    mSrcAddr = VarPtr(mSrcInt)
    
    #If Enable_Int Then
        
        #If Enable_PropertyGet Then
            
            mDestInt = 0
            mSrcInt = 12345
            mDestInt = IntMem(mSrcAddr)
            zAssert mDestInt, mSrcInt, "mDestInt = IntMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestInt = 0
            mSrcInt = 12345
            IntMem(mDestAddr) = mSrcInt
            zAssert mDestInt, mSrcInt, "IntMem(mDestAddr) = mSrcInt"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestInt = 0
            mSrcInt = 12345
            GetMem2 mSrcAddr, mDestInt
            zAssert mDestInt, mSrcInt, "GetMem2 mSrcAddr, mDestInt"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestInt = 0
            mSrcInt = 12345
            PutMem2 mDestAddr, mSrcInt
            zAssert mDestInt, mSrcInt, "PutMem2 mDestAddr, mSrcInt"
    
        #End If
        
        zSuccess "Test_IntAccess"
        
    #End If
        
End Sub

' Long memory access
' **********************

Public Sub Test_LongAccess()
    Dim mDestLong As Long, mSrcLong As Long

    mDestAddr = VarPtr(mDestLong)
    mSrcAddr = VarPtr(mSrcLong)
    
    #If Enable_Long Then
        
        #If Enable_PropertyGet Then
            
            mDestLong = 0
            mSrcLong = 123567890
            mDestLong = LongMem(mSrcAddr)
            zAssert mDestLong, mSrcLong, "mDestLong = LongMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestLong = 0
            mSrcLong = 123567890
            LongMem(mDestAddr) = mSrcLong
            zAssert mDestLong, mSrcLong, "LongMem(mDestAddr) = mSrcLong"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestLong = 0
            mSrcLong = 123567890
            GetMem4 mSrcAddr, mDestLong
            zAssert mDestLong, mSrcLong, "GetMem4 mSrcAddr, mDestLong"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestLong = 0
            mSrcLong = 123567890
            PutMem4 mDestAddr, mSrcLong
            zAssert mDestLong, mSrcLong, "PutMem4 mDestAddr, mSrcLong"
    
        #End If
        
        zSuccess "Test_LongAccess"
        
    #End If
        
End Sub

' Cur memory access
' **********************

Public Sub Test_CurAccess()
    Dim mDestCur As Currency, mSrcCur As Currency

    mDestAddr = VarPtr(mDestCur)
    mSrcAddr = VarPtr(mSrcCur)
    
    #If Enable_Cur Then
        
        #If Enable_PropertyGet Then
            
            mDestCur = 0
            mSrcCur = 12356789012345.6789@
            mDestCur = CurMem(mSrcAddr)
            zAssert mDestCur, mSrcCur, "mDestCur = CurMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestCur = 0
            mSrcCur = 12356789012345.6789@
            CurMem(mDestAddr) = mSrcCur
            zAssert mDestCur, mSrcCur, "CurMem(mDestAddr) = mSrcCur"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestCur = 0
            mSrcCur = 12356789012345.6789@
            GetMem8 mSrcAddr, mDestCur
            zAssert mDestCur, mSrcCur, "GetMem8 mSrcAddr, mDestCur"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestCur = 0
            mSrcCur = 12356789012345.6789@
            PutMem8 mDestAddr, mSrcCur
            zAssert mDestCur, mSrcCur, "PutMem8 mDestAddr, mSrcCur"
    
        #End If
        
        zSuccess "Test_CurAccess"
        
    #End If
        
End Sub

' Sng memory access
' **********************

Public Sub Test_SngAccess()
    Dim mDestSng As Single, mSrcSng As Single

    mDestAddr = VarPtr(mDestSng)
    mSrcAddr = VarPtr(mSrcSng)
    
    #If Enable_Sng Then
        
        #If Enable_PropertyGet Then
            
            mDestSng = 0
            mSrcSng = 123.4567
            mDestSng = SngMem(mSrcAddr)
            zAssert mDestSng, mSrcSng, "mDestSng = SngMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestSng = 0
            mSrcSng = 123.4567
            SngMem(mDestAddr) = mSrcSng
            zAssert mDestSng, mSrcSng, "SngMem(mDestAddr) = mSrcSng"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestSng = 0
            mSrcSng = 123.4567
            GetMem4_Sng mSrcAddr, mDestSng
            zAssert mDestSng, mSrcSng, "GetMem4_Sng mSrcAddr, mDestSng"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestSng = 0
            mSrcSng = 123.4567
            PutMem4_Sng mDestAddr, mSrcSng
            zAssert mDestSng, mSrcSng, "PutMem4_Sng mDestAddr, mSrcSng"
    
        #End If
        
        zSuccess "Test_SngAccess"
        
    #End If
        
End Sub

' Dbl memory access
' **********************

Public Sub Test_DblAccess()
    Dim mDestDbl As Double, mSrcDbl As Double

    mDestAddr = VarPtr(mDestDbl)
    mSrcAddr = VarPtr(mSrcDbl)
    
    #If Enable_Dbl Then
        
        #If Enable_PropertyGet Then
            
            mDestDbl = 0
            mSrcDbl = 12345678901.234
            mDestDbl = DblMem(mSrcAddr)
            zAssert mDestDbl, mSrcDbl, "mDestDbl = DblMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestDbl = 0
            mSrcDbl = 12345678901.234
            DblMem(mDestAddr) = mSrcDbl
            zAssert mDestDbl, mSrcDbl, "DblMem(mDestAddr) = mSrcDbl"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestDbl = 0
            mSrcDbl = 12345678901.234
            GetMem8_Dbl mSrcAddr, mDestDbl
            zAssert mDestDbl, mSrcDbl, "GetMem8_Dbl mSrcAddr, mDestDbl"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestDbl = 0
            mSrcDbl = 12345678901.234
            PutMem8_Dbl mDestAddr, mSrcDbl
            zAssert mDestDbl, mSrcDbl, "PutMem8_Dbl mDestAddr, mSrcDbl"
    
        #End If
        
        zSuccess "Test_DblAccess"
        
    #End If
        
End Sub

' Date memory access
' **********************

Public Sub Test_DateAccess()
    Dim mDestDate As Date, mSrcDate As Date

    mDestAddr = VarPtr(mDestDate)
    mSrcAddr = VarPtr(mSrcDate)
    
    #If Enable_Date Then
        
        #If Enable_PropertyGet Then
            
            mDestDate = 0
            mSrcDate = Now
            mDestDate = DateMem(mSrcAddr)
            zAssert mDestDate, mSrcDate, "mDestDate = DateMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestDate = 0
            mSrcDate = Now
            DateMem(mDestAddr) = mSrcDate
            zAssert mDestDate, mSrcDate, "DateMem(mDestAddr) = mSrcDate"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestDate = 0
            mSrcDate = Now
            GetMem8_Date mSrcAddr, mDestDate
            zAssert mDestDate, mSrcDate, "GetMem8_Date mSrcAddr, mDestDate"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestDate = 0
            mSrcDate = Now
            PutMem8_Date mDestAddr, mSrcDate
            zAssert mDestDate, mSrcDate, "PutMem8_Date mDestAddr, mSrcDate"
    
        #End If
        
        zSuccess "Test_DateAccess"
        
    #End If
        
End Sub

' LLong memory access
' **********************

Public Sub Test_LLongAccess()
    Dim mDestLLong As LongLong, mSrcLLong As LongLong

    mDestAddr = VarPtr(mDestLLong)
    mSrcAddr = VarPtr(mSrcLLong)
    
    #If Enable_LLong And (Win64 Or TWINBASIC) Then
        
        #If Enable_PropertyGet Then
            
            mDestLLong = 0
            mSrcLLong = 1234567890123456789^
            mDestLLong = LLongMem(mSrcAddr)
            zAssert mDestLLong, mSrcLLong, "mDestLLong = LLongMem(mSrcAddr)"
        
        #End If
        
        #If Enable_PropertyLet Then
            
            mDestLLong = 0
            mSrcLLong = 1234567890123456789^
            LLongMem(mDestAddr) = mSrcLLong
            zAssert mDestLLong, mSrcLLong, "LLongMem(mDestAddr) = mSrcLLong"
        
        #End If
    
        #If Enable_GetMem Then
        
            mDestLLong = 0
            mSrcLLong = 1234567890123456789^
            GetMem8_LLong mSrcAddr, mDestLLong
            zAssert mDestLLong, mSrcLLong, "GetMem8_LLong mSrcAddr, mDestLLong"
    
        #End If
        
        #If Enable_PutMem Then
    
            mDestLLong = 0
            mSrcLLong = 1234567890123456789^
            PutMem8_LLong mDestAddr, mSrcLLong
            zAssert mDestLLong, mSrcLLong, "PutMem8_LLong mDestAddr, mSrcLLong"
    
        #End If
        
        zSuccess "Test_LLongAccess"
        
    #End If
        
End Sub

Private Sub zAssert(Dest, src, Optional ByVal Text$)
    If Dest = src Then Exit Sub
    Debug.Print "*** "; Text; " ***  -> Dest = "; Dest; " ,  Src ="; src
    Debug.Assert Dest = src
End Sub

Private Sub zSuccess(ByVal Text$)
    Debug.Print "*** "; Text; " ***  -> successful"
End Sub
