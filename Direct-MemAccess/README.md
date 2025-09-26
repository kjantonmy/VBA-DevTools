# VBA Direct Memory Access
 * This module provides methods for accessing memory directly (low level access by memory address/pointer)


It works on either **`32/64-bit`**, and on any **`MS Office`** versions and platforms (**`Windows`** or **`Mac`**).  
It also works on **`TwinBasic`**.

## Notes:
* This module is intended for advanced programmers only who have in-depth knowledge of inner workings of VBA, also Win32 and OLE/COM automation
* Use with caution and at your own risk

## Compiler constants to disable unwanted groups of methods
If you need to disable any group of unused methods, the following compiler constants can be modified from `1` to `0`.  
Alternatively, the compiler constants can be commented out by inserting `'` at the beginning of the line.

* `#Const Enable_MoveMem = 1`
   * .
* `#Const Enable_GetMem = 1`
* `#Const Enable_PutMem = 1`
* `#Const Enable_PropertyGet = 1`
* `#Const Enable_PropertyLet = 1`
   * .
* `#Const Enable_Ptr = 1`
* `#Const Enable_Byte = 1`
* `#Const Enable_Int = 1`
* `#Const Enable_Long = 1`
* `#Const Enable_Cur = 1`
* `#Const Enable_Sng = 1`
* `#Const Enable_Dbl = 1`
* `#Const Enable_Date = 1`
* `#Const Enable_LLong = 1`


## Methods
 * Copy a block of memory bytes (correctly copies even if there is an overlap between destination and source)
   * `MoveMem`(ByVal DestAddr As LongPtr, ByVal SrcAddr As LongPtr, ByVal Length As LongPtr)

 * RetVal = _xxx_`Mem`(ByVal Addr As LongPtr)
    * RetVal = `PtrMem`(Addr)
    * RetVal = `ByteMem`(Addr)
    * RetVal = `IntMem`(Addr)
    * RetVal = `LongMem`(Addr)
    * RetVal = `CurMem`(Addr)
    * RetVal = `SngMem`(Addr)
    * RetVal = `DblMem`(Addr)
    * RetVal = `DateMem`(Addr)
    * RetVal = `LLongMem`(Addr)

 * _xxx_`Mem`(ByVal Addr As LongPtr) = NewVal
    * `PtrMem`(Addr) = NewVal
    * `ByteMem`(Addr) = NewVal
    * `IntMem`(Addr) = NewVal
    * `LongMem`(Addr) = NewVal
    * `CurMem`(Addr) = NewVal
    * `SngMem`(Addr) = NewVal
    * `DblMem`(Addr) = NewVal
    * `DateMem`(Addr) = NewVal
    * `LLongMem`(Addr) = NewVal

 * `GetMem`_xxx_(ByVal Addr As LongPtr, ByRef RetVal As _Type_)
    * `GetMemPtr`(Addr, ByRef RetVal As LongPtr)
    * `GetMem1`(Addr, ByRef RetVal As Byte)
    * `GetMem2`(Addr, ByRef RetVal As Integer)
    * `GetMem4`(Addr, ByRef RetVal As Long)
    * `GetMem8`(Addr, ByRef RetVal As Currency)
    * `GetMem4_Sng`(Addr, ByRef RetVal As Single)
    * `GetMem8_Dbl`(Addr, ByRef RetVal As Double)
    * `GetMem8_Date`(Addr, ByRef RetVal As Date)
    * `GetMem8_LLong`(Addr, ByRef RetVal As LongLong)

 * `PutMem`_xxx_(ByVal Addr As LongPtr, ByVal NewVal As _Type_)
    * `PutMemPtr`(Addr, NewVal As LongPtr)
    * `PutMem1`(Addr, NewVal As Byte)
    * `PutMem2`(Addr, NewVal As Integer)
    * `PutMem4`(Addr, NewVal As Long)
    * `PutMem8`(Addr, NewVal As Currency)
    * `PutMem4_Sng`(Addr, NewVal As Single)
    * `PutMem8_Dbl`(Addr, NewVal As Double)
    * `PutMem8_Date`(Addr, NewVal As Date)
    * `PutMem8_LLong`(Addr, NewVal As LongLong)
