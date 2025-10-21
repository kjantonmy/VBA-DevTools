[README.md](https://github.com/user-attachments/files/23021854/README.md)
# VBA Developer Tools
This package contains 1 class module(Nano_StopWatch.cls) and 2 standard modules(DirectMem.bas and Faster_StrQSort.bas)

### Nano_StopWatch.cls (Nano Seconds StopWatch)

* This class provides stop watch functionality with nano seconds and micro seconds precision.
* It is built around **_`QueryPerformanceCounter`_** API.


### DirectMem.bas (Direct Memory Access)
 * This module provides methods for accessing memory directly (low level access by memory address/pointer)

### Faster_StrQSort.bas (Faster String Quick Sort)
 * This module provides a method for sorting a flat String array (1D array).
 * It uses a tweaked version of Quick Sort by replacing String swaps with indirect String pointer swaps.
 * It also utilizes some additinal speed optimizations.

All of these modules works on either **`32/64-bit`**, and on any **`MS Office`** versions and platforms (**`Windows`** or **`Mac`**).  
It also works on **`TwinBasic`**.

