Attribute VB_Name = "DirectMem"
' Copyright Ken J. Anton 2025

''# VBA Direct Memory Access
'' * This module provides methods for accessing memory directly (low level access by memory address/pointer)


''It works on either **`32/64-bit`**, and on any **`MS Office`** versions and platforms (**`Windows`** or **`Mac`**).
''It also works on **`TwinBasic`**.

''## Notes:
''* This module is intended for advanced programmers only who have in-depth knowledge of inner workings of VBA, also Win32 and OLE/COM automation
''* Use with caution and at your own risk

''## Compiler constants to disable unwanted groups of methods
''If you need to disable any group of unused methods, the following compiler constants can be modified from `1` to `0`.
''Alternatively, the compiler constants can be commented out by inserting `'` at the beginning of the line.

''* `#Const Enable_MoveMem = 1`
''  * .
''* `#Const Enable_GetMem = 1`
''* `#Const Enable_PutMem = 1`
''* `#Const Enable_PropertyGet = 1`
''* `#Const Enable_PropertyLet = 1`
''  * .
''* `#Const Enable_Ptr = 1`
''* `#Const Enable_Byte = 1`
''* `#Const Enable_Int = 1`
''* `#Const Enable_Long = 1`
''* `#Const Enable_Cur = 1`
''* `#Const Enable_Sng = 1`
''* `#Const Enable_Dbl = 1`
''* `#Const Enable_Date = 1`
''* `#Const Enable_LLong = 1`


''## Methods
'' * Copy a block of memory bytes (correctly copies even if there is an overlap between destination and source)
''   * `MoveMem`(ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)

'' * RetVal = _xxx_`Mem`(ByVal Addr As LongPtr)
''    * RetVal = `PtrMem`(Addr)
''    * RetVal = `ByteMem`(Addr)
''    * RetVal = `IntMem`(Addr)
''    * RetVal = `LongMem`(Addr)
''    * RetVal = `CurMem`(Addr)
''    * RetVal = `SngMem`(Addr)
''    * RetVal = `DblMem`(Addr)
''    * RetVal = `DateMem`(Addr)
''    * RetVal = `LLongMem`(Addr)

'' * _xxx_`Mem`(ByVal Addr As LongPtr) = NewVal
''    * `PtrMem`(Addr) = NewVal
''    * `ByteMem`(Addr) = NewVal
''    * `IntMem`(Addr) = NewVal
''    * `LongMem`(Addr) = NewVal
''    * `CurMem`(Addr) = NewVal
''    * `SngMem`(Addr) = NewVal
''    * `DblMem`(Addr) = NewVal
''    * `DateMem`(Addr) = NewVal
''    * `LLongMem`(Addr) = NewVal

'' * `GetMem`_xxx_(ByVal Addr As LongPtr, ByRef RetVal As _Type_)
''    * `GetMemPtr`(Addr, ByRef RetVal As LongPtr)
''    * `GetMem1`(Addr, ByRef RetVal As Byte)
''    * `GetMem2`(Addr, ByRef RetVal As Integer)
''    * `GetMem4`(Addr, ByRef RetVal As Long)
''    * `GetMem8`(Addr, ByRef RetVal As Currency)
''    * `GetMem4_Sng`(Addr, ByRef RetVal As Single)
''    * `GetMem8_Dbl`(Addr, ByRef RetVal As Double)
''    * `GetMem8_Date`(Addr, ByRef RetVal As Date)
''    * `GetMem8_LLong`(Addr, ByRef RetVal As LongLong)

'' * `PutMem`_xxx_(ByVal Addr As LongPtr, ByVal NewVal As _Type_)
''    * `PutMemPtr`(Addr, NewVal As LongPtr)
''    * `PutMem1`(Addr, NewVal As Byte)
''    * `PutMem2`(Addr, NewVal As Integer)
''    * `PutMem4`(Addr, NewVal As Long)
''    * `PutMem8`(Addr, NewVal As Currency)
''    * `PutMem4_Sng`(Addr, NewVal As Single)
''    * `PutMem8_Dbl`(Addr, NewVal As Double)
''    * `PutMem8_Date`(Addr, NewVal As Date)
''    * `PutMem8_LLong`(Addr, NewVal As LongLong)

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

' Define LongPtr type for legacy VBA6

#If VBA7 = 0 And TWINBASIC = 0 Then
    ' This LongPtr type definition is for legacy 32-bit platform only.
    ' In case this type has been defined somewhere else in another module,
    '   this definition must be removed to prevent 'ambiguous type' error!
    
    Public Enum LongPtr ' By default, Enum type value is a Long type
        [_] ' Empty Enum
    End Enum
    
#End If

' Only 2 flags needed to be used

Public Enum SAFEARRAY_FeatureFlags
    FADF_AUTO = &H1
    'FADF_STATIC = &H2
    'FADF_EMBEDDED = &H4
    FADF_FIXEDSIZE = &H10
    'FADF_RECORD = &H20
    'FADF_HAVEIID = &H40
    'FADF_HAVEVARTYPE = &H80
    'FADF_BSTR = &H100
    'FADF_UNKNOWN = &H200
    'FADF_DISPATCH = &H400
    'FADF_VARIANT = &H800
    'FADF_RESERVED = &HF008
End Enum

' Packing together Data ptr and SafeArray Bound
Public Type SA_DataRange
    p As LongPtr
    Cnt As Long
    IndexLB As Long
End Type

Private mEmptySaDataRg As SA_DataRange

' Alternate SafeArray structure
Public Type SA_Type1
   Dims As Integer
   FeatureFlags As Integer
   eSize As Long
   LockCnt As Long
   DataRg As SA_DataRange
End Type

Private Type SA_PtrMem:     sa As SA_Type1:     m() As LongPtr:     End Type

#If Win64 Then
    Public Const PtrLen& = 8
#Else
    Public Const PtrLen& = 4
#End If

' Compiler Constants for different versions/platforms

#Const LegacyVBA6 = VBA6 And VBA7 = 0
#Const TwinBasic_or_VBA7 = TWINBASIC Or VBA7
#Const Mac_or_VBA6 = Mac Or LegacyVBA6

#Const MacVBA6 = Mac And LegacyVBA6
#Const MacVBA7 = Mac And VBA7
#Const Win64_NonMac_NonTwinBasic = Win64 And Mac = 0 And TWINBASIC = 0

' Declaring MoveMem for direct API call for Mac/TwinBasic or 32-bit version (except for 64-bit version of MS Office for windows)

#If Enable_MoveMem Then

    #If MacVBA7 Then
        Public Declare PtrSafe Sub MoveMem Lib "/usr/lib/libc.dylib" Alias "memmove" (ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)
    #ElseIf TWINBASIC Then
        Public Declare PtrSafe Sub MoveMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)
    #ElseIf MacVBA6 Then
        Public Declare Sub MoveMem Lib "/usr/lib/libc.dylib" Alias "memmove" (ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)
    #ElseIf LegacyVBA6 Then
        Public Declare Sub MoveMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)
    #End If

#End If

' Declaring xMoveByRef for direct API call for all versions. It is intended for internal use only

#If MacVBA7 Then
    Private Declare PtrSafe Sub xMoveByRef Lib "/usr/lib/libc.dylib" Alias "memmove" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
#ElseIf TwinBasic_or_VBA7 Then
    Private Declare PtrSafe Sub xMoveByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
#ElseIf MacVBA6 Then
    Private Declare  Sub xMoveByRef Lib "/usr/lib/libc.dylib" Alias "memmove" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
#ElseIf LegacyVBA6 Then
    Private Declare  Sub xMoveByRef Lib "kernel32" Alias "RtlMoveMemory" (ByRef DestRef As Any, ByRef SrcRef As Any, ByVal Length As LongPtr)
#End If

' Declaring GetMem/PutMem for direct API call on TwinBasic for Single, Double, Date and LongLong types

#If TWINBASIC Then

    #If Enable_GetMem Then
    
        #If Enable_Sng Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub GetMem4_Sng Lib "<hiddenmodule>" Alias "#5" (ByVal Address As LongPtr, ByRef RetVal As Single)
        #End If
        
        #If Enable_Dbl Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub GetMem8_Dbl Lib "<hiddenmodule>" Alias "#6" (ByVal Address As LongPtr, ByRef RetVal As Double)
        #End If

        #If Enable_Date Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub GetMem8_Date Lib "<hiddenmodule>" Alias "#6" (ByVal Address As LongPtr, ByRef RetVal As Date)
        #End If

        #If Enable_LLong Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub GetMem8_LLong Lib "<hiddenmodule>" Alias "#6" (ByVal Address As LongPtr, ByRef RetVal As LongLong)
        #End If
        
    #End If
    
    #If Enable_PutMem Then
    
        #If Enable_Sng Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub PutMem4_Sng Lib "<hiddenmodule>" Alias "#10" (ByVal Address As LongPtr, ByVal NewVal As Single)
        #End If

        #If Enable_Dbl Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub PutMem8_Dbl Lib "<hiddenmodule>" Alias "#11" (ByVal Address As LongPtr, ByVal NewVal As Double)
        #End If

        #If Enable_Date Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub PutMem8_Date Lib "<hiddenmodule>" Alias "#11" (ByVal Address As LongPtr, ByVal NewVal As Date)
        #End If

        #If Enable_LLong Then
            [PreserveSig(False), UseGetLastError(False), DLLStackCheck(False)]
            Public DeclareWide PtrSafe Sub PutMem8_LLong Lib "<hiddenmodule>" Alias "#11" (ByVal Address As LongPtr, ByVal NewVal As LongLong)
        #End If
        
    #End If

#End If

' SafeArray type definitions for each of enabled types

#If Enable_Byte Then
    Public Const ByteLen& = 1
    Private Type SA_ByteMem:   sa As SA_Type1:    m() As Byte:      End Type
#End If

#If Enable_Int Then
    Public Const IntLen& = 2
    Private Type SA_IntMem:    sa As SA_Type1:    m() As Integer:   End Type
#End If

#If Enable_Long Then
    Public Const LongLen& = 4
    Private Type SA_LongMem:   sa As SA_Type1:    m() As Long:      End Type
#End If

#If Enable_Cur Then
    Public Const CurLen& = 8
    Private Type SA_CurMem:    sa As SA_Type1:    m() As Currency:  End Type
#End If

#If Enable_Sng Then
    Public Const SngLen& = 4
    Private Type SA_SngMem:    sa As SA_Type1:    m() As Single:    End Type
#End If

#If Enable_Dbl Then
    Public Const DblLen& = 8
    Private Type SA_DblMem:    sa As SA_Type1:    m() As Double:    End Type
#End If

#If Enable_Date Then
    Public Const DateLen& = 8
    Private Type SA_DateMem:   sa As SA_Type1:    m() As Date:      End Type
#End If

#If TWINBASIC Or Win64 And Enable_LLong Then
    Public Const LLongLen& = 8
    Private Type SA_LLongMem:  sa As SA_Type1:    m() As LongLong:  End Type
#End If


#If Enable_MoveMem And Win64_NonMac_NonTwinBasic Then

' Definition of different sizes of memory block
' *** note: VBA limits fixed/static UDT type to less then 64k

    Private Type t2:     lo  As Byte:       hi  As Byte:        End Type
    Private Type t3:     lo2 As t2:         hi1 As Byte:        End Type
    Private Type t4:     lo  As t2:         hi  As t2:          End Type
    
    Private Type t5:     lo4 As t4:         hi1 As Byte:        End Type
    Private Type t6:     lo4 As t4:         hi2 As t2:          End Type
    Private Type t7:     lo6 As t6:         hi1 As Byte:        End Type
    
    Private Type t6_2:   lo6 As t6:         hi2 As t2:          End Type
    Private Type t7_1:   lo7 As t7:         hi1 As Byte:        End Type
    
    Private Type t16:    lo  As Currency:   hi  As Currency:    End Type
    Private Type t24:    lo16 As t16:       hi8 As Currency:    End Type
    Private Type t32:    lo As t16:         hi As t16:          End Type
    
    Private Type t40:    lo32 As t32:       hi8 As Currency:    End Type
    Private Type t48:    lo32 As t32:       hi16 As t16:        End Type
    Private Type t56:    lo32 As t32:       hi24 As t24:        End Type
    
    Private Type t64:    lo As t32:         hi As t32:          End Type
    Private Type t128:   lo As t64:         hi As t64:          End Type
    Private Type t3x64:  lo128 As t128:     hi64 As t64:        End Type

    Private Type t256:   lo As t128:        hi As t128:         End Type
    Private Type t512:   lo As t256:        hi As t256:         End Type
    Private Type t3x256: lo512 As t512:     hi256 As t256:      End Type
    Private Type t1k:    lo As t512:        hi As t512:         End Type
    
    Private Type t5x256: lo1k As t1k:       hi256 As t256:      End Type
    Private Type t6x256: lo1k As t1k:       hi512 As t512:      End Type
    Private Type t7x256: lo1k As t1k:       hi3x256 As t3x256:  End Type
    
    Private Type t2k:    lo As t1k:         hi As t1k:          End Type
    Private Type t4k:    lo As t2k:         hi As t2k:          End Type
    Private Type t6k:    lo4k As t4k:       hi2k As t2k:        End Type
    
    Private Type t8k:    lo As t4k:         hi As t4k:          End Type
    Private Type t10k:   lo8k As t8k:       hi2k As t2k:        End Type
    Private Type t12k:   lo8k As t8k:       hi4k As t4k:        End Type
    Private Type t14k:   lo8k As t8k:       hi6k As t6k:        End Type
    
    Private Type t16k:   lo As t8k:         hi As t8k:          End Type
    Private Type t32k:   lo As t16k:        hi As t16k:         End Type
    Private Type t48k:   lo32k As t32k:     hi16k As t16k:      End Type
    
    Private Type TmpMemType
        m1k As t1k:      m3x256 As t3x256:  m512 As t512:       m128 As t128:     m64 As t64
        m7 As t7:        m1 As Byte:        m6 As t6:           m2 As t2
    End Type

    Private Type SA_MemTbl
        sa As SA_Type1
        
        m1() As Byte:      m2() As Integer:     m3() As t3:
        m4() As Long:      m5() As t5:          m6() As t6:         m7() As t7
        
        m8() As Currency:  m16() As t16:        m24() As t24
        m32() As t32:      m40() As t40:        m48() As t48:       m56() As t56
        
        m64() As t64:      m128() As t128:      m3x64() As t3x64
        
        m256() As t256:    m512() As t512:      m3x256() As t3x256
        m1k() As t1k:      m5x256() As t5x256:  m6x256() As t6x256: m7x256() As t7x256
        
        m2k() As t2k:      m4k() As t4k:        m6k() As t6k
        m8k() As t8k:      m10k() As t10k:      m12k() As t12k:     m14k() As t14k
        
        m16k() As t16k:    m32k() As t32k:      m48k() As t48k
        
'--------
        s32() As String * 16:       s3x64() As String * 96
        
        s3x256() As String * 384:   s5x256() As String * 640:   s6x256() As String * 768:  s7x256() As String * 896
        
        s2k() As String * 1024:     s4k() As String * 2048:     s6k() As String * 3072
        s8k() As String * 4096:     s10k() As String * 5120:    s12k() As String * 6144:   s14k() As String * 7168
        s16k() As String * 8192:    s32k() As String * 16384:   s48k() As String * 24576
        
    End Type
    
' Definition of Split Indexing Table

Private Type SplitNdxType
        x1 As Byte: x2 As Byte: x3 As Byte
        topSplit As Byte
        split1 As Byte: split2 As Byte: split3 As Byte
        lastSplit As Byte
    End Type
    
    Private Type SplitTblType
        mv(0 To 255) As SplitNdxType
        mvo(0 To 255) As SplitNdxType
    End Type
    
#End If

' Ptr (LongPtr) memory access
' *******************************
#If Enable_Ptr Then

    #If Enable_PropertyGet Then

        Public Property Get PtrMem(ByVal Addr As LongPtr) As LongPtr
             #If Mac_or_VBA6 Then
                xMoveByRef PtrMem, ByVal Addr, PtrLen
            #ElseIf TWINBASIC Then
                GetMemPtr Addr, PtrMem
            #Else
                Static m As SA_PtrMem: If m.sa.Dims = 0 Then zInit_Link_SA m.sa, PtrLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                PtrMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
    
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let PtrMem(ByVal Addr As LongPtr, ByVal NewVal As LongPtr)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, PtrLen
            #ElseIf TWINBASIC Then
                PutMemPtr Addr, NewVal
            #Else
                Static m As SA_PtrMem: If m.sa.Dims = 0 Then zInit_Link_SA m.sa, PtrLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If

    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMemPtr(ByVal Addr As LongPtr, ByRef RetVal As LongPtr)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, PtrLen
            #Else
                Static m As SA_PtrMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, PtrLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMemPtr(ByVal Addr As LongPtr, ByVal NewVal As LongPtr)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, PtrLen
            #Else
                Static m As SA_PtrMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, PtrLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If

#End If
    
' Byte memory access
' **********************

#If Enable_Byte Then
    
    #If Enable_PropertyGet Then
    
        Public Property Get ByteMem(ByVal Addr As LongPtr) As Byte
            #If Mac_or_VBA6 Then
                xMoveByRef ByteMem, ByVal Addr, ByteLen
            #ElseIf TWINBASIC Then
                GetMem1 Addr, ByteMem
            #Else
                Static m As SA_ByteMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, ByteLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                ByteMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let ByteMem(ByVal Addr As LongPtr, ByVal NewVal As Byte)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, ByteLen
            #ElseIf TWINBASIC Then
                PutMem1 Addr, NewVal
            #Else
                Static m As SA_ByteMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, ByteLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
    
    #End If

    #If Enable_GetMem And TWINBASIC = 0 Then
    
        Public Sub GetMem1(ByVal Addr As LongPtr, ByRef RetVal As Byte)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, ByteLen
            #Else
                Static m As SA_ByteMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, ByteLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem1(ByVal Addr As LongPtr, ByVal NewVal As Byte)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, ByteLen
            #Else
                Static m As SA_ByteMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, ByteLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' Int memory access
' **********************

#If Enable_Int Then

    #If Enable_PropertyGet Then
    
        Public Property Get IntMem(ByVal Addr As LongPtr) As Integer
            #If Mac_or_VBA6 Then
                xMoveByRef IntMem, ByVal Addr, IntLen
            #ElseIf TWINBASIC Then
                GetMem2 Addr, IntMem
            #Else
                Static m As SA_IntMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, IntLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                IntMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let IntMem(ByVal Addr As LongPtr, ByVal NewVal As Integer)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, IntLen
            #ElseIf TWINBASIC Then
                PutMem2 Addr, NewVal
            #Else
                Static m As SA_IntMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, IntLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then
    
        Public Sub GetMem2(ByVal Addr As LongPtr, ByRef RetVal As Integer)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, IntLen
            #Else
                Static m As SA_IntMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, IntLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem2(ByVal Addr As LongPtr, ByVal NewVal As Integer)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, IntLen
            #Else
                Static m As SA_IntMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, IntLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' Long memory access
' **********************
    
#If Enable_Long Then
    
    #If Enable_PropertyGet Then
        
        Public Property Get LongMem(ByVal Addr As LongPtr) As Long
            #If Mac_or_VBA6 Then
                xMoveByRef LongMem, ByVal Addr, LongLen
            #ElseIf TWINBASIC Then
                GetMem4 Addr, LongMem
            #Else
                Static m As SA_LongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                LongMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let LongMem(ByVal Addr As LongPtr, ByVal NewVal As Long)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, LongLen
            #ElseIf TWINBASIC Then
                PutMem4 Addr, NewVal
            #Else
                Static m As SA_LongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem4(ByVal Addr As LongPtr, ByRef RetVal As Long)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, LongLen
            #Else
                Static m As SA_LongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem4(ByVal Addr As LongPtr, ByVal NewVal As Long)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, LongLen
            #Else
                Static m As SA_LongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' Cur memory access
' **********************
    
#If Enable_Cur Then
        
    #If Enable_PropertyGet Then
        
        Public Property Get CurMem(ByVal Addr As LongPtr) As Currency
            #If Mac_or_VBA6 Then
                xMoveByRef CurMem, ByVal Addr, CurLen
            #ElseIf TWINBASIC Then
                GetMem8 Addr, CurMem
            #Else
                Static m As SA_CurMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, CurLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                CurMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let CurMem(ByVal Addr As LongPtr, ByVal NewVal As Currency)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, CurLen
            #ElseIf TWINBASIC Then
                PutMem8 Addr, NewVal
            #Else
                Static m As SA_CurMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, CurLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem8(ByVal Addr As LongPtr, ByRef RetVal As Currency)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, CurLen
            #Else
                Static m As SA_CurMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, CurLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem8(ByVal Addr As LongPtr, ByVal NewVal As Currency)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, CurLen
            #Else
                Static m As SA_CurMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, CurLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' Sng memory access
' **********************
    
#If Enable_Sng Then
        
    #If Enable_PropertyGet Then
        
        Public Property Get SngMem(ByVal Addr As LongPtr) As Single
            #If Mac_or_VBA6 Then
                xMoveByRef SngMem, ByVal Addr, SngLen
            #ElseIf TWINBASIC Then
                GetMem4_Sng Addr, SngMem
            #Else
                Static m As SA_SngMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, SngLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                SngMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let SngMem(ByVal Addr As LongPtr, ByVal NewVal As Single)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, SngLen
            #ElseIf TWINBASIC Then
                PutMem4_Sng Addr, NewVal
            #Else
                Static m As SA_SngMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, SngLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem4_Sng(ByVal Addr As LongPtr, ByRef RetVal As Single)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, SngLen
            #Else
                Static m As SA_SngMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, SngLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem4_Sng(ByVal Addr As LongPtr, ByVal NewVal As Single)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, SngLen
            #Else
                Static m As SA_SngMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, SngLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub
        
    #End If
    
#End If
    
' Dbl memory access
' **********************
    
#If Enable_Dbl Then
        
    #If Enable_PropertyGet Then
        
        Public Property Get DblMem(ByVal Addr As LongPtr) As Double
            #If Mac_or_VBA6 Then
                xMoveByRef DblMem, ByVal Addr, DblLen
            #ElseIf TWINBASIC Then
                GetMem8_Dbl Addr, DblMem
            #Else
                Static m As SA_DblMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DblLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                DblMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let DblMem(ByVal Addr As LongPtr, ByVal NewVal As Double)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, DblLen
            #ElseIf TWINBASIC Then
                PutMem8_Dbl Addr, NewVal
            #Else
                Static m As SA_DblMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DblLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem8_Dbl(ByVal Addr As LongPtr, ByRef RetVal As Double)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, DblLen
            #Else
                Static m As SA_DblMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DblLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem8_Dbl(ByVal Addr As LongPtr, ByVal NewVal As Double)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, DblLen
            #Else
                Static m As SA_DblMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DblLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' Date memory access
' **********************
    
#If Enable_Date Then
        
    #If Enable_PropertyGet Then
        
        Public Property Get DateMem(ByVal Addr As LongPtr) As Date
            #If Mac_or_VBA6 Then
                xMoveByRef DateMem, ByVal Addr, DateLen
            #ElseIf TWINBASIC Then
                GetMem8_Date Addr, DateMem
            #Else
                Static m As SA_DateMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DateLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                DateMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let DateMem(ByVal Addr As LongPtr, ByVal NewVal As Date)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, DateLen
            #ElseIf TWINBASIC Then
                PutMem8_Date Addr, NewVal
            #Else
                Static m As SA_DateMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DateLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
    
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem8_Date(ByVal Addr As LongPtr, ByRef RetVal As Date)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, DateLen
            #Else
                Static m As SA_DateMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DateLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem8_Date(ByVal Addr As LongPtr, ByVal NewVal As Date)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, DateLen
            #Else
                Static m As SA_DateMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, DateLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If
    
#End If
    
' LLong memory access
' ************************
    
#If Enable_LLong And (Win64 Or TWINBASIC) Then
    
    #If Enable_PropertyGet Then
        
        Public Property Get LLongMem(ByVal Addr As LongPtr) As LongLong
            #If Mac_or_VBA6 Then
                xMoveByRef LLongMem, ByVal Addr, LLongLen
            #ElseIf TWINBASIC Then
                GetMem8_LLong Addr, LLongMem
            #Else
                Static m As SA_LLongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LLongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                LLongMem = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
        
    #End If
    
    #If Enable_PropertyLet Then
        
        Public Property Let LLongMem(ByVal Addr As LongPtr, ByVal NewVal As LongLong)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, LLongLen
            #ElseIf TWINBASIC Then
                PutMem8_LLong Addr, NewVal
            #Else
                Static m As SA_LLongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LLongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Property
    
    #End If
    
    #If Enable_GetMem And TWINBASIC = 0 Then

        Public Sub GetMem8_LLong(ByVal Addr As LongPtr, ByRef RetVal As LongLong)
            #If Mac_or_VBA6 Then
                xMoveByRef RetVal, ByVal Addr, LLongLen
            #Else
                Static m As SA_LLongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LLongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                RetVal = m.m(0)
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub
        
    #End If
    
    #If Enable_PutMem And TWINBASIC = 0 Then

        Public Sub PutMem8_LLong(ByVal Addr As LongPtr, ByVal NewVal As LongLong)
            #If Mac_or_VBA6 Then
                xMoveByRef ByVal Addr, NewVal, LLongLen
            #Else
                Static m As SA_LLongMem:   If m.sa.Dims = 0 Then zInit_Link_SA m.sa, LLongLen
                m.sa.DataRg.Cnt = 1:  m.sa.DataRg.p = Addr
                m.m(0) = NewVal
                m.sa.DataRg = mEmptySaDataRg
            #End If
        End Sub

    #End If

#End If

        
' Faster version of "RtlMoveMemory" native API for 64-bit Windows only
' *********************************************************************
' Note: this method uses 'GoTo' and 'On ... GoTo' statements extensively for speed optimization


#If Enable_MoveMem And Win64_NonMac_NonTwinBasic Then

    Public Sub MoveMem(ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)
        Const n1k& = 1024, n2k& = 2 * n1k, n4k& = 4 * n1k, n8k& = 8 * n1k, n16k& = 16 * n1k, n32k& = 32 * n1k, n48k& = 48 * n1k, n64k& = 64 * n1k
        Const Beyond_32k_Range& = -n32k, Beyond_Threshhold& = -&H4000000, LongMaxVal& = &H7FFFFFFF
        
        #If Win64 Then
            Const PtrMaxVal^ = &H7FFFFFFFFFFFFFFF^
        #Else
            Const PtrMaxVal& = LongMaxVal
        #End If
        
        Static dst As SA_MemTbl, src As SA_MemTbl, mSplitTbl As SplitTblType, tmp  As TmpMemType
        
        Dim nLen&, nSplit1&, nSplit2&, nCnt&, Gap As LongPtr, sx As SplitNdxType
    
    '********
        If dst.sa.Dims = 0 Then zInit_StaticData mSplitTbl, dst, src
    
        Gap = DestAddr - SrcAddr
        dst.sa.DataRg.Cnt = 1:      src.sa.DataRg.Cnt = 1
        
        If (Length - 1) And -8 Then GoTo mv_beyond_8_range ' (Length < 1 or > 8)
        
        dst.sa.DataRg.p = DestAddr:     src.sa.DataRg.p = SrcAddr
        
        On CLng(Length) GoTo mv_1z, mv_2z, mv_3gz, mv_4z, mv_5gz, mv_6gz, mv_7gz, mv_8z
            
    '********
mv_beyond_8_range:
        If Length And Beyond_32k_Range Then GoTo mv_beyond_32k_range ' (Length < 0 or >= 32k)
        
        nLen = CLng(Length)
        If (Gap And PtrMaxVal) < nLen Then GoTo mv_overlap
    
        dst.sa.DataRg.p = DestAddr:     src.sa.DataRg.p = SrcAddr
        
mv_below_64k:
        If nLen > 255 Then GoTo mv_start_below_64k
        
mv_start_below_256:
        sx = mSplitTbl.mv(nLen)
        On sx.x1 GoTo zzPreExit, _
            mv_1z, mv_2z, mv_3z, mv_4z, mv_5z, mv_6z, mv_7z, _
            mv_8z, mv_16z, mv_24z, mv_32z, mv_40z, mv_48z, mv_56z, _
            mv_8, mv_16, mv_24, mv_32, mv_40, mv_48, mv_56, _
            mv_64z, mv_128z, mv_3x64z, _
            mv_64, mv_128, mv_3x64
    
mv_below_64:
        dst.sa.DataRg.p = dst.sa.DataRg.p + sx.split1:    src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        On sx.x2 GoTo zzPreExit, _
            mv_1z, mv_2z, mv_3z, mv_4z, mv_5z, mv_6z, mv_7z, _
            mv_8z, mv_16z, mv_24z, mv_32z, mv_40z, mv_48z, mv_56z, _
            mv_8, mv_16, mv_24, mv_32, mv_40, mv_48, mv_56
    
mv_below_8:
        dst.sa.DataRg.p = dst.sa.DataRg.p + sx.split2:    src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        On sx.x3 GoTo zzPreExit, _
            mv_1z, mv_2z, mv_3z, mv_4z, mv_5z, mv_6z, mv_7z
    
    '********
mv_start_below_64k:
        sx = mSplitTbl.mv(nLen \ 256)
        nLen = nLen And 255 ' reminder lower 8-bit split
        On sx.x1 GoTo zzPreExit, _
            mv_256, mv_512, mv_3x256, mv_1k, mv_5x256, mv_6x256, mv_7x256, _
            mv_2kZ, mv_4kZ, mv_6kZ, mv_8kZ, mv_10kZ, mv_12kZ, mv_14kZ, _
            mv_2k, mv_4k, mv_6k, mv_8k, mv_10k, mv_12k, mv_14k, _
            mv_16kZ, mv_32kZ, mv_48kZ, _
            mv_16k, mv_32k, mv_48k
        
mv_below_16k:
        dst.sa.DataRg.p = dst.sa.DataRg.p + sx.split1 * 256&:  src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        On sx.x2 GoTo zzPreExit, _
            mv_256, mv_512, mv_3x256, mv_1k, mv_5x256, mv_6x256, mv_7x256, _
            mv_2kZ, mv_4kZ, mv_6kZ, mv_8kZ, mv_10kZ, mv_12kZ, mv_14kZ, _
            mv_2k, mv_4k, mv_6k, mv_8k, mv_10k, mv_12k, mv_14k
        
mv_below_2k:
        dst.sa.DataRg.p = dst.sa.DataRg.p + sx.split2 * 256&:  src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        On sx.x3 GoTo zzPreExit, _
            mv_256z, mv_512z, mv_3x256z, mv_1kZ, mv_5x256z, mv_6x256z, mv_7x256z, _
            mv_256, mv_512, mv_3x256, mv_1k, mv_5x256, mv_6x256, mv_7x256
        
mv_below_256:
        If nLen = 0 Then GoTo zzPreExit
        dst.sa.DataRg.p = dst.sa.DataRg.p + sx.lastSplit * 256&:  src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        GoTo mv_start_below_256
    
mv_1z:      dst.m1(0) = src.m1(0):  GoTo zzPreExit
    
mv_3gz:     If (Gap And PtrMaxVal) >= Length Then GoTo mv_3z
            dst.m3(0).hi1 = src.m3(0).hi1
mv_2z:      dst.m2(0) = src.m2(0):  GoTo zzPreExit
    
mv_3z:      dst.m3(0) = src.m3(0):  GoTo zzPreExit
    
mv_5gz:     If (Gap And PtrMaxVal) >= Length Then GoTo mv_5z
            dst.m5(0).hi1 = src.m5(0).hi1
mv_4z:      dst.m4(0) = src.m4(0):  GoTo zzPreExit
    
mv_5z:      dst.m5(0) = src.m5(0):  GoTo zzPreExit
mv_6z:      dst.m6(0) = src.m6(0):  GoTo zzPreExit
mv_7z:      dst.m7(0) = src.m7(0):  GoTo zzPreExit
mv_8z:      dst.m8(0) = src.m8(0):  GoTo zzPreExit
    
mv_6gz:     If (Gap And PtrMaxVal) >= Length Then GoTo mv_6z
            tmp.m6 = src.m6(0): dst.m6(0) = tmp.m6: GoTo zzPreExit
                
mv_7gz:     If (Gap And PtrMaxVal) >= Length Then GoTo mv_7z
            tmp.m7 = src.m7(0): dst.m7(0) = tmp.m7: GoTo zzPreExit
    
    '--------
mv_16z:     dst.m16(0) = src.m16(0):    GoTo zzPreExit
mv_24z:     dst.m24(0) = src.m24(0):    GoTo zzPreExit
mv_32z:     dst.m32(0) = src.m32(0):    GoTo zzPreExit
mv_40z:     dst.m40(0) = src.m40(0):    GoTo zzPreExit
mv_48z:     dst.m48(0) = src.m48(0):    GoTo zzPreExit
mv_56z:     dst.m56(0) = src.m56(0):    GoTo zzPreExit
    
mv_64z:         dst.m64(0) = src.m64(0):        GoTo zzPreExit
mv_128z:        dst.m128(0) = src.m128(0):      GoTo zzPreExit
mv_3x64z:       dst.m3x64(0) = src.m3x64(0):    GoTo zzPreExit
    
mv_256z:        dst.m256(0) = src.m256(0):      GoTo mv_below_256
mv_512z:        dst.m512(0) = src.m512(0):      GoTo mv_below_256
mv_3x256z:      dst.m3x256(0) = src.m3x256(0):  GoTo mv_below_256
mv_1kZ:         dst.m1k(0) = src.m1k(0):        GoTo mv_below_256
mv_5x256z:      dst.m5x256(0) = src.m5x256(0):  GoTo mv_below_256
mv_6x256z:      dst.m6x256(0) = src.m6x256(0):  GoTo mv_below_256
mv_7x256z:      dst.m7x256(0) = src.m7x256(0):  GoTo mv_below_256
    
mv_2kZ:         dst.m2k(0) = src.m2k(0):        GoTo mv_below_256
mv_4kZ:         dst.m4k(0) = src.m4k(0):        GoTo mv_below_256
mv_6kZ:         dst.m6k(0) = src.m6k(0):        GoTo mv_below_256
mv_8kZ:         dst.m8k(0) = src.m8k(0):        GoTo mv_below_256
mv_10kZ:        dst.m10k(0) = src.m10k(0):      GoTo mv_below_256
mv_12kZ:        dst.m12k(0) = src.m12k(0):      GoTo mv_below_256
mv_14kZ:        dst.m14k(0) = src.m14k(0):      GoTo mv_below_256
    
mv_16kZ:        dst.m16k(0) = src.m16k(0):      GoTo mv_below_256
mv_32kZ:        dst.m32k(0) = src.m32k(0):      GoTo mv_below_256
mv_48kZ:        dst.m48k(0) = src.m48k(0):      GoTo mv_below_256
    
    '--------
mv_8:      dst.m8(0) = src.m8(0):       GoTo mv_below_8
mv_16:      dst.m16(0) = src.m16(0):    GoTo mv_below_8
mv_24:      dst.m24(0) = src.m24(0):    GoTo mv_below_8
mv_32:      dst.m32(0) = src.m32(0):    GoTo mv_below_8
mv_40:      dst.m40(0) = src.m40(0):    GoTo mv_below_8
mv_48:      dst.m48(0) = src.m48(0):    GoTo mv_below_8
mv_56:      dst.m56(0) = src.m56(0):    GoTo mv_below_8
    
mv_64:          dst.m64(0) = src.m64(0):        GoTo mv_below_64
mv_128:         dst.m128(0) = src.m128(0):      GoTo mv_below_64
mv_3x64:        dst.m3x64(0) = src.m3x64(0):    GoTo mv_below_64
    
mv_256:         dst.m256(0) = src.m256(0):      GoTo mv_below_256
mv_512:         dst.m512(0) = src.m512(0):      GoTo mv_below_256
mv_3x256:       dst.m3x256(0) = src.m3x256(0):  GoTo mv_below_256
mv_1k:          dst.m1k(0) = src.m1k(0):        GoTo mv_below_256
mv_5x256:       dst.m5x256(0) = src.m5x256(0):  GoTo mv_below_256
mv_6x256:       dst.m6x256(0) = src.m6x256(0):  GoTo mv_below_256
mv_7x256:       dst.m7x256(0) = src.m7x256(0):  GoTo mv_below_256
    
mv_2k:          dst.m2k(0) = src.m2k(0):        GoTo mv_below_2k
mv_4k:          dst.m4k(0) = src.m4k(0):        GoTo mv_below_2k
mv_6k:          dst.m6k(0) = src.m6k(0):        GoTo mv_below_2k
mv_8k:          dst.m8k(0) = src.m8k(0):        GoTo mv_below_2k
mv_10k:         dst.m10k(0) = src.m10k(0):      GoTo mv_below_2k
mv_12k:         dst.m12k(0) = src.m12k(0):      GoTo mv_below_2k
mv_14k:         dst.m14k(0) = src.m14k(0):      GoTo mv_below_2k
    
mv_16k:         dst.m16k(0) = src.m16k(0):      GoTo mv_below_16k
mv_32k:         dst.m32k(0) = src.m32k(0):      GoTo mv_below_16k
mv_48k:         dst.m48k(0) = src.m48k(0):      GoTo mv_below_16k
    
    '********
mv_overlap:     ' assumes nLen >= 9  and < 64k
        If Gap = 0 Then GoTo zzPreExit
    
        sx = mSplitTbl.mvo(nLen And 255)
        
        dst.sa.DataRg.p = DestAddr + nLen - sx.split1:   src.sa.DataRg.p = dst.sa.DataRg.p - Gap
                                            
        If Gap >= sx.split1 Then On sx.x1 GoTo mvo_above255, _
            mvo_3x64, mvo_128, mvo_64, _
            mvo_56, mvo_48, mvo_40, mvo_32, mvo_24, mvo_16, mvo_8, _
            mvo_7, mvo_6, mvo_5, mvo_4, mvo_3, mvo_2, mvo_1
        
        If Gap < sx.lastSplit Then On sx.x1 GoTo mvo_above255, _
            mvo_s3x64, mvo_s128, mvo_s64, _
            mvo_s56, mvo_s48, mvo_s40, mvo_s32, mvo_s24, mvo_s16, mvo_s8, _
            mvo_s7, mvo_s6, mvo_s5, mvo_4, mvo_s3, mvo_2, mvo_1
        
        ' Gap >= sx.lastSplit
        On sx.x1 GoTo mvo_above255, _
            mvo_g3x64, mvo_g128, mvo_g64, _
            mvo_s56, mvo_s48, mvo_s40, mvo_s32, mvo_s24, mvo_s16, mvo_s8, _
            mvo_s7, mvo_s6, mvo_s5, mvo_4, mvo_s3, mvo_2, mvo_1
        
    '--------
mvo_s7:     tmp.m7 = src.m7(0): dst.m7(0) = tmp.m7:     GoTo mvo_above7
    
mvo_7:      dst.m7(0) = src.m7(0):                                     GoTo mvo_above7
    
mvo_s6:     tmp.m6 = src.m6(0): dst.m6(0) = tmp.m6:     GoTo mvo_above7
    
mvo_6:      dst.m6(0) = src.m6(0):      GoTo mvo_above7
    
mvo_5:      dst.m5(0) = src.m5(0):      GoTo mvo_above7
    
mvo_s5:     dst.m5(0).hi1 = src.m5(0).hi1
mvo_4:      dst.m4(0) = src.m4(0):      GoTo mvo_above7
    
mvo_3:      dst.m3(0) = src.m3(0):      GoTo mvo_above7
    
mvo_s3:     dst.m3(0).hi1 = src.m3(0).hi1
mvo_2:      dst.m2(0) = src.m2(0):      GoTo mvo_above7
    
mvo_1:      dst.m1(0) = src.m1(0) ':    GoTo mvo_above7
    
    '****************
mvo_above7:
    
        dst.sa.DataRg.p = dst.sa.DataRg.p - sx.split2:  src.sa.DataRg.p = dst.sa.DataRg.p - Gap
    
        If Gap < sx.topSplit Then On sx.x2 GoTo mvo_above255, _
            mvo_s3x64, mvo_s128, mvo_s64, _
            mvo_s56, mvo_s48, mvo_s40, mvo_s32, mvo_s24, mvo_s16, mvo_s8
    
        If Gap >= sx.split2 Then On sx.x2 GoTo mvo_above255, _
            mvo_3x64, mvo_128, mvo_64, _
            mvo_56, mvo_48, mvo_40, mvo_32, mvo_24, mvo_16, mvo_8
        
        ' Gap < sx.split2
        On sx.x2 GoTo mvo_above255, _
            mvo_g3x64, mvo_g128, mvo_g64, _
            mvo_s56, mvo_s48, mvo_s40, mvo_s32, mvo_s24, mvo_s16, mvo_s8
                                            
    '--------
mvo_56:         dst.m64(0).hi.hi.lo = src.m64(0).hi.hi.lo
mvo_48:         dst.m64(0).hi.lo = src.m64(0).hi.lo
                      dst.m32(0) = src.m32(0):    GoTo mvo_above63
    
mvo_40:         dst.m64(0).hi.lo.lo = src.m64(0).hi.lo.lo
mvo_32:         dst.m32(0) = src.m32(0):    GoTo mvo_above63
    
mvo_24:         dst.m32(0).hi.lo = src.m32(0).hi.lo
mvo_16:         dst.m16(0) = src.m16(0):    GoTo mvo_above63
    
    '--------
mvo_s56:        dst.m64(0).hi.hi.lo = src.m64(0).hi.hi.lo
mvo_s48:        dst.m64(0).hi.lo.hi = src.m64(0).hi.lo.hi
mvo_s40:        dst.m64(0).hi.lo.lo = src.m64(0).hi.lo.lo
    
mvo_s32:        dst.s32(0) = src.s32(0):    GoTo mvo_above63
    
mvo_s24:        dst.m32(0).hi.lo = src.m32(0).hi.lo
mvo_s16:        dst.m16(0).hi = src.m16(0).hi
    
mvo_s8:
mvo_8:          dst.m8(0) = src.m8(0) ':    GoTo mvo_above63
    
    '********
mvo_above63:
        If nLen < 64 Then GoTo zzPreExit
        If sx.split3 = 0 Then GoTo mvo_above255
        
        dst.sa.DataRg.p = dst.sa.DataRg.p - sx.split3:  src.sa.DataRg.p = dst.sa.DataRg.p - Gap
    
        If Gap < 64 Then On sx.x3 GoTo mvo_above255, mvo_s3x64, mvo_s128, mvo_s64
        If Gap >= sx.split3 Then On sx.x3 GoTo mvo_above255, mvo_3x64, mvo_128, mvo_64
        
        ' Gap < sx.split3
        On sx.x3 GoTo mvo_above255, mvo_g3x64, mvo_g128, mvo_g64
        
    '--------
mvo_3x64:       dst.m256(0).hi.lo = src.m256(0).hi.lo
mvo_128:        dst.m128(0) = src.m128(0):      GoTo mvo_above255
    
mvo_g3x64:      dst.m256(0).hi.lo = src.m256(0).hi.lo
mvo_g128:       dst.m128(0).hi = src.m128(0).hi
mvo_g64:
mvo_64:         dst.m64(0) = src.m64(0):        GoTo mvo_above255
    
mvo_s3x64:      dst.s3x64(0) = src.s3x64(0):    GoTo mvo_above255
mvo_s128:       tmp.m128 = src.m128(0): dst.m128(0) = tmp.m128:     GoTo mvo_above255
mvo_s64:        tmp.m64 = src.m64(0):   dst.m64(0) = tmp.m64 ':     GoTo mvo_above255
    
    '********
mvo_above255:
        If nLen < 256 Then GoTo zzPreExit
        sx = mSplitTbl.mvo((nLen \ 256)) ' assumes nLen <= 48k
        nSplit1 = sx.split1 * 256&
        dst.sa.DataRg.p = dst.sa.DataRg.p - nSplit1:    src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        
        If Gap < sx.lastSplit * 256 Then On sx.x1 GoTo zzPreExit, _
            mvo_s3x16k, mvo_s2x16k, mvo_s1x16k, _
            mvo_s7x2k, mvo_s6x2k, mvo_s5x2k, mvo_s4x2k, mvo_s3x2k, mvo_s2x2k, mvo_s1x2k, _
            mvo_s7x256, mvo_s6x256, mvo_s5x256, mvo_s4x256, mvo_s3x256, mvo_s2x256, mvo_s1x256
        
        If Gap < nSplit1 Then On sx.x1 GoTo zzPreExit, _
            mvo_g3x16k, mvo_g2x16k, mvo_g1x16k, _
            mvo_g7x2k, mvo_g6x2k, mvo_g5x2k, mvo_g4x2k, mvo_g3x2k, mvo_g2x2k, mvo_g1x2k, _
            mvo_g7x256, mvo_g6x256, mvo_g5x256, mvo_g4x256, mvo_g3x256, mvo_g2x256, mvo_g1x256
        
        ' Gap >= nSplit1
        On sx.x1 GoTo zzPreExit, _
            mvo_3x16k, mvo_2x16k, mvo_1x16k, _
            mvo_7x2k, mvo_6x2k, mvo_5x2k, mvo_4x2k, mvo_3x2k, mvo_2x2k, mvo_1x2k, _
            mvo_7x256, mvo_6x256, mvo_5x256, mvo_4x256, mvo_3x256, mvo_2x256, mvo_1x256
    
    '--------
mvo_s7x256:     dst.s7x256(0) = src.s7x256(0):              GoTo mvo_aboveEq_2k
mvo_s6x256:     dst.s6x256(0) = src.s6x256(0):              GoTo mvo_aboveEq_2k
mvo_s5x256:     dst.s5x256(0) = src.s5x256(0):              GoTo mvo_aboveEq_2k
mvo_s4x256:     tmp.m1k = src.m1k(0):           dst.m1k(0) = tmp.m1k:           GoTo mvo_aboveEq_2k
mvo_s3x256:     tmp.m3x256 = src.m3x256(0):     dst.m3x256(0) = tmp.m3x256:     GoTo mvo_aboveEq_2k
mvo_s2x256:     tmp.m3x256.lo512 = src.m512(0): dst.m512(0) = tmp.m3x256.lo512: GoTo mvo_aboveEq_2k
mvo_s1x256:     tmp.m3x256.hi256 = src.m256(0): dst.m256(0) = tmp.m3x256.hi256: GoTo mvo_aboveEq_2k
    
    '--------
mvo_g7x256:     dst.m2k(0).hi.hi.lo = src.m2k(0).hi.hi.lo
mvo_g6x256:     dst.m2k(0).hi.lo.hi = src.m2k(0).hi.lo.hi
mvo_g5x256:     dst.m2k(0).hi.lo.lo = src.m2k(0).hi.lo.lo
    
mvo_g4x256:     tmp.m1k = src.m1k(0): dst.m1k(0) = tmp.m1k: GoTo mvo_aboveEq_2k
    
mvo_g3x256:     dst.m1k(0).hi.lo = src.m1k(0).hi.lo
mvo_g2x256:     dst.m512(0).hi = src.m512(0).hi
    
mvo_g1x256:
mvo_1x256:      dst.m256(0) = src.m256(0):      GoTo mvo_aboveEq_2k
    
mvo_7x256:      dst.m7x256(0) = src.m7x256(0):  GoTo mvo_aboveEq_2k
    
mvo_6x256:      dst.m2k(0).hi.lo = src.m2k(0).hi.lo
                dst.m1k(0) = src.m1k(0):        GoTo mvo_aboveEq_2k
    
mvo_5x256:      dst.m2k(0).hi.lo.lo = src.m2k(0).hi.lo.lo
mvo_4x256:      dst.m1k(0) = src.m1k(0):        GoTo mvo_aboveEq_2k
    
mvo_3x256:      dst.m3x256(0) = src.m3x256(0):  GoTo mvo_aboveEq_2k
mvo_2x256:      dst.m512(0) = src.m512(0):      GoTo mvo_aboveEq_2k
    
    '--------
mvo_aboveEq_2k:
        If sx.split2 = 0 Then GoTo zzPreExit
        nSplit1 = sx.split2 * 256&
        dst.sa.DataRg.p = dst.sa.DataRg.p - nSplit1:    src.sa.DataRg.p = dst.sa.DataRg.p - Gap
    
        If Gap < sx.topSplit * 256& Then On sx.x2 GoTo zzPreExit, _
            mvo_s3x16k, mvo_s2x16k, mvo_s1x16k, _
            mvo_s7x2k, mvo_s6x2k, mvo_s5x2k, mvo_s4x2k, mvo_s3x2k, mvo_s2x2k, mvo_s1x2k
                                            
        If Gap < nSplit1 Then On sx.x2 GoTo zzPreExit, _
            mvo_g3x16k, mvo_g2x16k, mvo_g1x16k, _
            mvo_g7x2k, mvo_g6x2k, mvo_g5x2k, mvo_g4x2k, mvo_g3x2k, mvo_g2x2k, mvo_g1x2k
    
        ' Gap >= nSplit1
        On sx.x2 GoTo zzPreExit, _
            mvo_g3x16k, mvo_g2x16k, mvo_g1x16k, _
            mvo_7x2k, mvo_6x2k, mvo_5x2k, mvo_4x2k, mvo_3x2k, mvo_2x2k, mvo_1x2k
        
    '--------
mvo_s7x2k:      dst.s14k(0) = src.s14k(0):      GoTo mvo_aboveEq_16k
mvo_s6x2k:      dst.s12k(0) = src.s12k(0):      GoTo mvo_aboveEq_16k
mvo_s5x2k:      dst.s10k(0) = src.s10k(0):      GoTo mvo_aboveEq_16k
mvo_s4x2k:      dst.s8k(0) = src.s8k(0):        GoTo mvo_aboveEq_16k
mvo_s3x2k:      dst.s6k(0) = src.s6k(0):        GoTo mvo_aboveEq_16k
mvo_s2x2k:      dst.s4k(0) = src.s4k(0):        GoTo mvo_aboveEq_16k
mvo_s1x2k:      dst.s2k(0) = src.s2k(0):        GoTo mvo_aboveEq_16k
    
mvo_g7x2k:      dst.m16k(0).hi.hi.lo = src.m16k(0).hi.hi.lo
mvo_g6x2k:      dst.m16k(0).hi.lo.hi = src.m16k(0).hi.lo.hi
mvo_g5x2k:      dst.m16k(0).hi.lo.lo = src.m16k(0).hi.lo.lo
mvo_g4x2k:      dst.m8k(0).hi.hi = src.m8k(0).hi.hi
mvo_g3x2k:      dst.m8k(0).hi.lo = src.m8k(0).hi.lo
mvo_g2x2k:      dst.m4k(0).hi = src.m4k(0).hi
mvo_g1x2k:      dst.m2k(0) = src.m2k(0):                  GoTo mvo_aboveEq_16k
    
    '--------
mvo_7x2k:       dst.m14k(0) = src.m14k(0):      GoTo mvo_aboveEq_16k
    
mvo_6x2k:       dst.m16k(0).hi.lo = src.m16k(0).hi.lo
                dst.m8k(0) = src.m8k(0):        GoTo mvo_aboveEq_16k
    
mvo_5x2k:       dst.m16k(0).hi.lo.lo = src.m16k(0).hi.lo.lo
mvo_4x2k:       dst.m8k(0) = src.m8k(0):        GoTo mvo_aboveEq_16k
    
mvo_3x2k:       dst.m8k(0).hi.lo = src.m8k(0).hi.lo
mvo_2x2k:       dst.m4k(0) = src.m4k(0):        GoTo mvo_aboveEq_16k
    
mvo_1x2k:       dst.m2k(0) = src.m2k(0):        GoTo mvo_aboveEq_16k
    
    '--------
mvo_aboveEq_16k:
        If sx.split3 = 0 Then GoTo zzPreExit
        nSplit1 = sx.split3 * 256&
        dst.sa.DataRg.p = dst.sa.DataRg.p - nSplit1:    src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        
        If Gap < n16k Then
            If Gap >= n2k Then GoTo mvo_2k_4k_8k_loop
            On sx.x3 GoTo zzPreExit, mvo_s3x16k, mvo_s2x16k, mvo_s1x16k
        End If
            
        If Gap < nSplit1 Then On sx.x3 GoTo zzPreExit, mvo_g3x16k, mvo_g2x16k, mvo_g1x16k
        
        ' Gap >= nSplit1
        On sx.x3 GoTo zzPreExit, mvo_3x16k, mvo_2x16k, mvo_1x16k
        
    '--------
mvo_s3x16k:      If Gap >= n2k Then GoTo mvo_2k_4k_8k_loop Else dst.s48k(0) = src.s48k(0):    GoTo zzPreExit
mvo_s2x16k:      If Gap >= n2k Then GoTo mvo_2k_4k_8k_loop Else dst.s32k(0) = src.s32k(0):    GoTo zzPreExit
mvo_s1x16k:      If Gap >= n2k Then GoTo mvo_2k_4k_8k_loop Else dst.s16k(0) = src.s16k(0):    GoTo zzPreExit
    
mvo_g3x16k:      dst.m48k(0).hi16k = src.m48k(0).hi16k
mvo_g2x16k:      dst.m32k(0).hi = src.m32k(0).hi
mvo_1x16k:
mvo_g1x16k:      dst.m16k(0) = src.m16k(0):     GoTo zzPreExit
    
mvo_3x16k:       dst.m48k(0) = src.m48k(0):     GoTo zzPreExit
mvo_2x16k:       dst.m32k(0) = src.m32k(0):     GoTo zzPreExit
    
    '--------
mvo_2k_4k_8k_loop:
    
        If nLen < n16k Then GoTo zzPreExit
        nSplit2 = mSplitTbl.mv(CLng(Gap) \ 256).topSplit * 256&
        nCnt = nSplit1 \ nSplit2
        dst.sa.DataRg.Cnt = nCnt:    src.sa.DataRg.Cnt = nCnt:    dst.sa.eSize = nSplit2:    src.sa.eSize = nSplit2
        On nSplit2 \ n2k GoTo mvo_2k_loop, mvo_4k_loop, mvo_4k_loop, mvo_8k_loop
        Debug.Assert Gap >= n2k And Gap <= n8k
        GoTo zzPreExit
    
mvo_8k_loop:        Do While nCnt > 0:  nCnt = nCnt - 1:    dst.m8k(nCnt) = src.m8k(nCnt):    Loop:   GoTo zzPreExit2
mvo_4k_loop:        Do While nCnt > 0:  nCnt = nCnt - 1:    dst.m4k(nCnt) = src.m4k(nCnt):    Loop:   GoTo zzPreExit2
mvo_2k_loop:        Do While nCnt > 0:  nCnt = nCnt - 1:    dst.m2k(nCnt) = src.m2k(nCnt):    Loop:   GoTo zzPreExit2
    
    '********
mv_beyond_32k_range:
        If Gap = 0 Then GoTo zzPreExit
        If Length And Beyond_Threshhold Then GoTo mv_beyond_threshhold
        
        nLen = CLng(Length)
        If (Gap And PtrMaxVal) >= Length Then GoTo mv_aboveEq_32k
        If Length \ Gap <= 32 And Length < n64k Then GoTo mv_overlap
        GoTo mv_nativeAPI
        
mv_aboveEq_32k:
        dst.sa.DataRg.p = DestAddr:     src.sa.DataRg.p = SrcAddr
        If nLen < n64k Then GoTo mv_start_below_64k
        
        nSplit2 = nLen \ n32k
        nSplit1 = nLen And Beyond_32k_Range
        nLen = nLen And &H7FFF
        
        dst.sa.DataRg.Cnt = nSplit2:   src.sa.DataRg.Cnt = nSplit2:   dst.sa.eSize = n32k:   src.sa.eSize = n32k
    
        For nCnt = 0 To nSplit2 - 1:    dst.m32k(nCnt) = src.m32k(nCnt):    Next
        
        If nLen = 0 Then GoTo zzPreExit2
        
        dst.sa.DataRg.Cnt = 1:   src.sa.DataRg.Cnt = 1:     dst.sa.eSize = 0:   src.sa.eSize = 0
        dst.sa.DataRg.p = DestAddr + nSplit1:   src.sa.DataRg.p = dst.sa.DataRg.p - Gap
        
        GoTo mv_below_64k
        
    '********
mv_beyond_threshhold:
        If Length <= 0 Then GoTo zzPreExit
        
mv_nativeAPI:
        xMoveByRef ByVal DestAddr, ByVal SrcAddr, Length
        GoTo zzPreExit
        
    '********
zzPreExit2:
        dst.sa.eSize = 0:     src.sa.eSize = 0
        
zzPreExit:
        dst.sa.DataRg = mEmptySaDataRg:     src.sa.DataRg = mEmptySaDataRg
    End Sub
    
    Private Sub zInit_SplitTbl(ByRef SplitTbl As SplitTblType)
        Dim i&, ExitIndex&, IndexBase&, Index&, sx As SplitNdxType
            
        For i = 0 To 255
            ExitIndex = 1
            
            With SplitTbl.mv(i)
               
                Index = i And 7
                .x3 = IIf(Index, Index + 1, ExitIndex)
                .split3 = IIf(Index, Index, 0)
                
                Index = (i \ 8) And 7
                .x2 = IIf(Index, Index + 8 + ((.x3 > 1) And 7), .x3)
                .split2 = IIf(Index, Index * 8, .split3)
                
                Index = (i \ 64) And 7
                .x1 = IIf(Index, Index + 22 + ((.x2 > 1) And 3), .x2)
                .split1 = IIf(Index, Index * 64, .split2)
                
                .lastSplit = IIf(.split3, .split3, IIf(.split2, .split2, .split1))
                
                If i And 1 Then .topSplit = 1
                If i And 2 Then .topSplit = 2
                If i And 4 Then .topSplit = 4
                If i And 8 Then .topSplit = 8
                If i And 16 Then .topSplit = 16
                If i And 32 Then .topSplit = 32
                If i And 64 Then .topSplit = 64
                If i And 128 Then .topSplit = 128
                
            End With
    
            With SplitTbl.mvo(i)
               
                Index = (i \ 64) And 7
                .x3 = IIf(Index, 5 - Index, ExitIndex)
                .split3 = IIf(Index, Index * 64, 0)
                
                Index = (i \ 8) And 7
                .x2 = IIf(Index, 12 - Index, .x3)
                .split2 = IIf(Index, Index * 8, .split3)
                
                Index = i And 7
                .x1 = IIf(Index, 19 - Index, .x2)
                .split1 = IIf(Index, Index, .split2)
                
                Select Case .split1
                Case Is < 8: .lastSplit = 1:   .topSplit = IIf(.split2 < 64, 8, 64)
                Case Is < 64: .lastSplit = 8:  .topSplit = 64
                Case Else: .lastSplit = 64
                End Select
            End With
    
        Next
    End Sub
    
    Private Sub zInit_StaticData(ByRef SplitTbl As SplitTblType, ByRef sa1 As SA_MemTbl, ByRef sa2 As SA_MemTbl)
        Dim Cnt&
        
        Cnt = (LenB(sa1) - LenB(sa1.sa)) \ PtrLen
        zInit_Link_SA sa1.sa, 0, Cnt
        zInit_Link_SA sa2.sa, 0, Cnt
        
        zInit_SplitTbl SplitTbl
    
    End Sub
    
#End If

Private Function zInit_SA_Type1(ByRef sa As SA_Type1, Optional ByVal eSize&) As LongPtr
    With sa
        .Dims = 1
        .eSize = eSize
        .LockCnt = 1
        .FeatureFlags = FADF_AUTO Or FADF_FIXEDSIZE 'this will prevent from garbage collection
    End With
    zInit_SA_Type1 = VarPtr(sa)
End Function

Private Sub zInit_Link_SA(ByRef sa As SA_Type1, Optional ByVal eSize&, Optional ByVal MemArr_Cnt& = 1)
    Dim AddrOf_SA As LongPtr, i&
    
    Static m As SA_PtrMem
    If m.sa.Dims = 0 Then
        AddrOf_SA = zInit_SA_Type1(m.sa, PtrLen)
        xMoveByRef ByVal AddrOf_SA + LenB(m.sa), AddrOf_SA, PtrLen
    End If
    
    If MemArr_Cnt <= 0 Then MemArr_Cnt = 1
    
    AddrOf_SA = zInit_SA_Type1(sa, eSize)
    m.sa.DataRg.Cnt = MemArr_Cnt
    m.sa.DataRg.p = AddrOf_SA + LenB(sa)
    
    Do
        MemArr_Cnt = MemArr_Cnt - 1
        m.m(MemArr_Cnt) = AddrOf_SA
    Loop Until MemArr_Cnt <= 0
    
    m.sa.DataRg = mEmptySaDataRg
End Sub
