VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Nano_StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright Ken J. Anton 2025

' Nano_StopWatch    - class module
' *******************

'This is a class module built around `QueryPerformanceCounter` API.

'It provides methods for measuring elapsed time of any task/operation easily.
'When `StartWatch` is invoked, it automatically calculate/calibrate the elapsed overhead time.
'The net elapsed time sometime may fluctuate a little bit due to background activities of the operating system.

'It works on either `32/64-bit`, and on any `MS Office` versions and platforms (`Windows` or `Mac`).
'It also works on `TwinBasic`.

'## Methods

'* `StartWatch`
'* `StopWatch`

'*  `Elapsed`         (optional `TimeUnit`)  -> returns net elapsed time as formatted with time unit
'*  `Elapsed_Num`(optional `TimeUnit`)  -> returns net elapsed time as numeric value(double)
'    * _"`TimeUnit`"_ can be preset using `Default_TimeUnit` property

'*  `Default_TimeUnit` - property

'    * `"TimeUnit"` can be any of:  `auto`, `tick`, `s`/`sec`/`second`, `ms`/`milli`, `us`/`micro`, `ns`/`nano`
'    * `"TimeUnit"` can be suffixed with `s` (i.e. "ticks", "secs", "micros" and so on)

'*  `Elapsed_Overhead`  -> returns the overhead in ticks as numeric value(double)

'*  `Frequency`            ->  returns clock ticks per second as formatted(`"#,##0"`)
'*  `Frequency_Num`   ->  returns Frequency as numeric value(double)


Option Explicit

Const kMicroSign$ = "u"  ' can be replaced with literal micro sign `µ`, if the character set of your locale is Latin-1

Const kMicroSecSuffix = " " & kMicroSign & "s"

Const kUnitList$ = "~auto, ~tick, 3~n, 3~nano, 2~u, 2~" & kMicroSign & ", 2~micro, 1~m, 1~milli, 0~s, 0~sec, 0~second,", kTicks$ = "ticks"
Const kFlag_Started& = 1, kFlag_Stopped& = 2, kFlag_Elapsed& = 3

Private mRawStart@, mRawStop@, mRawFrequency@
Private mOverheadTicks#, mNetTicks#, mGrossTicks#, mFrequencyTicks#
Private mIsCalibrating As Boolean, mFlags&, mFactors#(0 To 4), mUnits$(0 To 4), mDefault_TimeUnit$

#If Mac Then

    #If VBA7 Then
        Private Declare PtrSafe Function QueryPerformanceCounter _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/mso20.framework/mso20" (Counter@) As Long
        
        Private Declare PtrSafe Function QueryPerformanceFrequency _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/mso20.framework/mso20" (Frequency@) As Long
    #Else
        Private Declare Function QueryPerformanceCounter _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/mso20.framework/mso20" (Counter@) As Long
        
        Private Declare Function QueryPerformanceFrequency _
            Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/mso20.framework/mso20" (Frequency@) As Long
    #End If

#Else 'Win

    #If VBA7 Then
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (Counter@) As Long
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (Frequency@) As Long
    #Else
        Private Declare Function QueryPerformanceCounter Lib "kernel32" (Counter@) As Long
        Private Declare Function QueryPerformanceFrequency Lib "kernel32" (Frequency@) As Long
    #End If

#End If

' *** Primary methods

Public Sub StartWatch()
    If Not mIsCalibrating Then zCalibrate_Overhead
    mFlags = kFlag_Started
    DoEvents ' yield to any pending events (mostly screen updating)
    QueryPerformanceCounter mRawStart
End Sub

Public Sub StopWatch()
    QueryPerformanceCounter mRawStop
    zCalc_IfStarted
End Sub

Public Function Elapsed$(Optional ByVal TimeUnit$)
    QueryPerformanceCounter mRawStop
    zCalc_IfStarted
    Elapsed = zConvertTime(mNetTicks, TimeUnit, True)
End Function

Public Function Elapsed_Num#(Optional ByVal TimeUnit$)
    QueryPerformanceCounter mRawStop
    zCalc_IfStarted
    Elapsed_Num = zConvertTime(mNetTicks, TimeUnit)
End Function

' *** Secondary methods

Public Function Elapsed_Overhead#()
    Elapsed_Overhead = mOverheadTicks
End Function

Public Function Frequency$()
    Frequency = Format(mFrequencyTicks, "#,##0")
End Function

Public Function Frequency_Num#()
    Frequency_Num = mFrequencyTicks
End Function

' *** Properties Get/Set

Public Property Get Default_TimeUnit$()
    Default_TimeUnit = mDefault_TimeUnit
End Property

Public Property Let Default_TimeUnit(ByVal TimeUnit$)
    If zFind_TimeUnit(TimeUnit) > 0 Then mDefault_TimeUnit = TimeUnit
End Property

' *** Internal methods

Private Sub zCalc_IfStarted()
    If mFlags = kFlag_Started Then
        mGrossTicks = zMul_10_000(mRawStop - mRawStart)
        mNetTicks = mGrossTicks - mOverheadTicks
        If mNetTicks < 0 Then mNetTicks = 0
        mFlags = kFlag_Elapsed
    End If
End Sub

Private Function zFind_TimeUnit&(ByVal TimeUnit$)
    Dim nLen&
    
    TimeUnit = Trim$(TimeUnit)
    
    ' use default if empty
    If Len(TimeUnit) = 0 And Len(mDefault_TimeUnit) <> 0 Then TimeUnit = mDefault_TimeUnit
    nLen = Len(TimeUnit)
    
    ' remove suffix 's' if present
    If nLen > 1 Then
        If StrComp(Mid$(TimeUnit, nLen, 1), "s", vbTextCompare) = 0 Then nLen = nLen - 1
    End If
    
    zFind_TimeUnit = InStr(1, kUnitList, "~" & Left$(TimeUnit, nLen) & ",", vbTextCompare)
End Function

Private Function zConvertTime(ByVal Ticks#, Optional ByVal TimeUnit$, Optional ByVal OptFmt As Boolean) As Variant
    Dim p&, Exp3&
    
    p = zFind_TimeUnit(TimeUnit)
    
    If p <= 1 Then
        zConvertTime = zAutoFmt(zConvertToSec(Ticks))
    ElseIf p < 10 Then
        zConvertTime = Ticks
        If OptFmt Then zConvertTime = zFmtNumber(Ticks, " ticks")
    Else
        Exp3 = Mid$(kUnitList, p - 1, 1)
        zConvertTime = zConvertToSec(Ticks) * (1000 ^ Exp3)
        If OptFmt Then zConvertTime = zFmtNumber(zConvertTime, mUnits(4 - Exp3))
    End If
End Function

Private Function zAutoFmt$(ByVal Seconds#)
    Dim Exp3#
    
    If Seconds = 0 Then
        zAutoFmt = "0 ticks"
        Exit Function
    ElseIf Seconds < 0 Then
        zAutoFmt = "-" & zAutoFmt(-Seconds)
        Exit Function
    End If
    
    Exp3 = (Log(Seconds) / Log(1000)) + 4
    If Exp3 < 2 Then Exp3 = 1
    If Exp3 > 4 Then Exp3 = 4
    Exp3 = Left$(Exp3, 1) ' removes fraction if any
    
    zAutoFmt = zFmtNumber(Seconds * mFactors(Exp3), mUnits(Exp3))
End Function

Private Function zFmtNumber$(ByVal Value#, Optional ByVal Unit$)
    Dim Exp3#, nLen&, nTrim&, nSep&, tx$, dgt$, dValue#, dDelta#
    
    If Value = 0 Then
        tx = "0"
        GoTo zzAppend_Unit
    ElseIf Value < 0 Then
        zFmtNumber = "-" & zFmtNumber(-Value)
        Exit Function
    End If
    
    ' trim repetitive digits
    tx = Format(Value, "#0.##################")
    nSep = InStr(tx, ".")
    dgt$ = Right$(tx, 3)
    nLen = Len(tx)
    nTrim = nLen
    
    If nSep > 0 And InStr("333~667~001", dgt) Then
        Mid$(dgt$, 3, 1) = Left$(dgt, 1)
        nTrim = nLen - 3
        
        Do While nTrim >= nSep + 3
            If Mid$(tx, nTrim - 2, 3) <> dgt Then Exit Do
            nTrim = nTrim - 3
        Loop
        
        dgt = Left$(dgt, 1)
        Do While nTrim > nSep
            If Mid$(tx, nTrim, 1) <> dgt Then Exit Do
            nTrim = nTrim - 1
        Loop
        
        dValue = Left$(tx, nTrim)
        If dgt = "0" Then
            dDelta = Value - dValue
            If dDelta <> 0 Then
                If (Log(dDelta) / Log(10)) >= -6 Then dValue = Value
            End If
        Else
            dgt = dgt & "···" & dgt
        End If
    Else
        dValue = Value
    End If
    
    tx = Format(dValue, "#,##0.###`###`###`###`###`###")
    
    ' trim trailing separators
    nLen = Len(tx)
    nTrim = nLen
    
    Do While Mid$(tx, nTrim, 1) = "`"
        nTrim = nTrim - 1
    Loop
    
    If Len(dgt) > 3 Then
        tx = Left$(tx, nTrim) & dgt
    Else
        If Mid$(tx, nTrim, 1) = "." Then nTrim = nTrim - 1
        If nTrim < nLen Then tx = Left$(tx, nTrim)
    End If
    
zzAppend_Unit:
    If LenB(Unit) Then
        If Left$(Unit, 1) <> " " Then Unit = " " & Unit
    End If
    
    zFmtNumber = tx & Unit
End Function

Private Sub zStartStop(ByVal self As Nano_StopWatch)
    self.StartWatch
    self.StopWatch
End Sub

Private Sub zCalibrate_Overhead()
    Dim i&, n&, e&, eArr&(), eMin&, eMax&, nStat&(), eMaxFreq&, Runs&
    Const kCalibration_Runs& = 128, kCalibration_TblSize& = 128
    
    If mIsCalibrating Then Exit Sub
    mIsCalibrating = True
    
    Runs = kCalibration_Runs
    If Runs < 1 Then Runs = 1
    ReDim eArr(1 To Runs)
    
    ' tally raw stats
    eMin = 2 ^ 30
    
    For i = 1 To Runs
        zStartStop Me
        e = CLng(mGrossTicks)
        eArr(i) = e
        If e < eMin Then eMin = e
    Next
                        
    eMax = eMin + kCalibration_TblSize - 1
    ReDim nStat(eMin To eMax)
    
    ' tally frequency stats
    For i = 1 To Runs
        e = eArr(i)
        If e > eMax Then
            e = eMax
        ElseIf e < eMin Then
            e = eMin
        End If
        nStat(e) = nStat(e) + 1
    Next
    
    ' find the most frequent overhead
    For i = eMin To eMax
        n = nStat(i)
        If n > eMaxFreq Then
            eMaxFreq = n
            e = i
        End If
    Next
    
    mOverheadTicks = e
    mIsCalibrating = False
End Sub

Private Function zMul_10_000#(ByVal Value@)
    zMul_10_000 = CDbl(Value) * 10000
End Function

Private Function zConvertToSec#(ByVal Ticks#)
    zConvertToSec = Ticks / mFrequencyTicks
End Function

Private Sub zInit_FmtTbl()
    Dim n&, u As Variant
    
    n = 4
    For Each u In Array(" sec", " ms", kMicroSecSuffix, " ns")
        mFactors(n) = 1000 ^ (4 - n)
        mUnits(n) = u
        n = n - 1
    Next
    mFactors(0) = 1
    mUnits(0) = kTicks
End Sub

Private Sub Class_Initialize()
    Dim Dummy@
    QueryPerformanceFrequency mRawFrequency
    mFrequencyTicks = zMul_10_000(mRawFrequency)
    QueryPerformanceCounter Dummy
    zCalibrate_Overhead
    
    zInit_FmtTbl
    mDefault_TimeUnit = "tick"
End Sub
