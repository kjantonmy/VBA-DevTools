# Nano_StopWatch

This is a class module built around **_`QueryPerformanceCounter`_** API.

* It provides methods for measuring elapsed time of any task/operation easily.
* When `StartWatch` is invoked, it automatically calculate/calibrate the elapsed overhead time.
* The net elapsed time sometime may fluctuate a little bit due to background activities of the operating system or other running apps.

It works on either **`32/64-bit`**, and on any **`MS Office`** versions and platforms (**`Windows`** or **`Mac`**).  
It also works on **`TwinBasic`**.

## Methods

* `StartWatch`
* `StopWatch`

*  `Elapsed`    (optional _`TimeUnit`_)  -> returns net elapsed time as formatted with time unit
*  `Elapsed_Num`(optional _`TimeUnit`_)  -> returns net elapsed time as numeric value(double)
    * _"`TimeUnit`"_ can be preset using `Default_TimeUnit` property

*  `Default_TimeUnit` - property

    * _`"TimeUnit"`_ can be any of:  `auto`, `tick`, `s`/`sec`/`second`, `ms`/`milli`, `µs`/`us`/`micro`, `ns`/`nano`
    * _`"TimeUnit"`_ can be suffixed with `s` (i.e. "ticks", "secs", "micros" and so on)

*  `Elapsed_Overhead`  -> returns the overhead in ticks as numeric value(double)

*  `Frequency`       ->  returns clock ticks per second as formatted(`"#,##0"`)
*  `Frequency_Num`   ->  returns Frequency as numeric value(double)

## Notes on using micro sign `µ` as in `µs` time unit for Latin-1 character set
If you prefer to use the `µ` instead of `u` micro sign, you can modify the const `kMicroSign` definition at the beginning of the class module.
