Attribute VB_Name = "TestModule_Nano_StopWatch"
' Copyright Ken J. Anton 2025

Option Explicit

Public Sub Test_Nano_StopWatch()
    Dim i&, a() As Long
    Const nElements& = 10 ^ 7
    
    With New Nano_StopWatch
        zDebugPrint
        zDebugPrint "Frequency: " & .Frequency
        zDebugPrint "Overhead: " & .Elapsed_Overhead
        zDebugPrint
        
        .StartWatch
        .StopWatch
        
        zDebugPrint "*** Do nothing"
        zDebugPrint "Elapsed: " & .Elapsed
        zDebugPrint "Elapsed('auto'): " & .Elapsed("auto")
        zDebugPrint
        
        i = Asc("A")

        .StartWatch
            i = Asc("A")
        .StopWatch
        
        zDebugPrint "*** i = Asc('A')"
        zDebugPrint "Elapsed: " & .Elapsed
        zDebugPrint "Elapsed('auto'): " & .Elapsed("auto")
        zDebugPrint
        
        Dir "c:\*.*"

        .StartWatch
            Dir "c:\*.*"
        .StopWatch
        
        zDebugPrint "*** dir 'c:\*.*'"
            
        zDebugPrint "Elapsed('auto'): " & .Elapsed("auto")
        zDebugPrint "Elapsed('tick'): " & .Elapsed("tick")
        zDebugPrint "Elapsed('ns'): " & .Elapsed("ns")
'        zDebugPrint "Elapsed('µs'): " & .Elapsed("µs")
        zDebugPrint "Elapsed('us'): " & .Elapsed("us")
        zDebugPrint "Elapsed('ms'): " & .Elapsed("ms")
        zDebugPrint "Elapsed('sec'): " & .Elapsed("sec")
        zDebugPrint
        
        ReDim a(1 To nElements)
        
        .StartWatch
            For i = 1 To nElements
                a(i) = i
            Next
        .StopWatch
        
        zDebugPrint "*** Fill array of " & Format(nElements, "#,##0") & " elements"
            
        zDebugPrint
        zDebugPrint "Elapsed: " & .Elapsed
        zDebugPrint "Elapsed_Num: " & .Elapsed_Num
        zDebugPrint "Elapsed('auto'): " & .Elapsed("auto")

        .Default_TimeUnit = "ms"
        zDebugPrint
        zDebugPrint "*** Default_TimeUnit = 'ms'"
        zDebugPrint "Elapsed: " & .Elapsed
        zDebugPrint "Elapsed_Num: " & .Elapsed_Num

'        .Default_TimeUnit = "µs"
        .Default_TimeUnit = "us"
        zDebugPrint
        zDebugPrint "*** Default_TimeUnit = 'us'"
        zDebugPrint "Elapsed: " & .Elapsed
        zDebugPrint "Elapsed_Num: " & .Elapsed_Num

    End With
    
End Sub

Private Sub zDebugPrint(Optional ByVal Txt$)
    #if TwinBasic Then
        Console.Writeline Txt
    #else
        Debug.Print Txt
    #end if
End Sub