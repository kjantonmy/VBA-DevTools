Attribute VB_Name = "TestModule_Nano_StopWatch"
' Copyright Ken J. Anton 2025

Option Explicit

Public Sub Test_Nano_StopWatch()
    Dim i&, a() As Long
    Const nElements& = 10 ^ 7
    
    With New Nano_StopWatch
        Debug.Print
        Debug.Print "Frequency: "; .Frequency
        Debug.Print "Overhead: "; .Elapsed_Overhead
        Debug.Print
        
        .StartWatch
        .StopWatch
        
        Debug.Print "*** Do nothing"
        Debug.Print "Elapsed: "; .Elapsed
        Debug.Print "Elapsed('auto'): "; .Elapsed("auto")
        Debug.Print
        
        .StartWatch
            i = Asc("A")
        .StopWatch
        
        Debug.Print "*** i = Asc('A')"
        Debug.Print "Elapsed: "; .Elapsed
        Debug.Print "Elapsed('auto'): "; .Elapsed("auto")
        Debug.Print
        
        .StartWatch
            Dir "c:\*.*"
        .StopWatch
        
        Debug.Print "*** dir 'c:\*.*'"
            
        Debug.Print "Elapsed('auto'): "; .Elapsed("auto")
        Debug.Print "Elapsed('tick'): "; .Elapsed("tick")
        Debug.Print "Elapsed('ns'): "; .Elapsed("ns")
'        Debug.Print "Elapsed('µs'): "; .Elapsed("µs")
        Debug.Print "Elapsed('us'): "; .Elapsed("us")
        Debug.Print "Elapsed('ms'): "; .Elapsed("ms")
        Debug.Print "Elapsed('sec'): "; .Elapsed("sec")
        Debug.Print
        
        ReDim a(1 To nElements)
        
        .StartWatch
            For i = 1 To nElements
                a(i) = i
            Next
        .StopWatch
        
        Debug.Print "*** Fill array of "; Format(nElements, "#,##0"); " elements"
            
        Debug.Print
        Debug.Print "Elapsed: "; .Elapsed
        Debug.Print "Elapsed_Num: "; .Elapsed_Num
        Debug.Print "Elapsed('auto'): "; .Elapsed("auto")

        .Default_TimeUnit = "ms"
        Debug.Print
        Debug.Print "*** Default_TimeUnit = 'ms'"
        Debug.Print "Elapsed: "; .Elapsed
        Debug.Print "Elapsed_Num: "; .Elapsed_Num

'        .Default_TimeUnit = "µs"
        .Default_TimeUnit = "us"
        Debug.Print
        Debug.Print "*** Default_TimeUnit = 'us'"
        Debug.Print "Elapsed: "; .Elapsed
        Debug.Print "Elapsed_Num: "; .Elapsed_Num

    End With
    
End Sub
