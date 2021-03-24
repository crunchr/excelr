# excelr

Excelerated .xlsx generation that can create a basic .xlsx file with virtually 
no memory (~10MB), and in very little time (~1s for 1.5 million cells).

## Motivation

This is useful when you have the following requirements...

* You want to create a simple excel file without any bells or whistles, i.e.
  no fonts, no colours, freeze panes etc. a single sheet
* You are working with a large dataset, and you want to use as little memory
  as possible (for example if you are providing an excel export from a web
  api)

Normally in these circumstances JSON or CSV are good choices, but for our use
case these are problematic. Our users want to just open the downloaded file in
excel, so JSON is not really an option. CSV files can be opened in excel, but
since it is primarily a text format problems can arise with different locales.

Therefore we implemented this small library that can be used to generate a 
basic excel file extremely quickly (~1s for 1.5 million cells) and using very
little memory (~10MB). Since it uses a temporary file to prevent having to
materialize the whole file in memory, this speed is dependent on the type of
hard disk in use, but when working with SSDs this is extremely fast.

## Benchmark

Here we compare the generation of an excel file from various sized datasets. We
look at pandas, openpyxl (using the optimized write_only mode) and excelr. The
tests look at wall time and memory usage. We can see that both openpyxl and
excelr perform very well from a memory perspective, while excelr outperforms
openpyxl when looking at run time.

Of course the pandas variant has to first create a DataFrame before it can 
export to excel, which explains the poor performance in terms of memory and 
run time.

![Alt text](/img/MB.png?raw=true "Memory usage")
![Alt text](/img/seconds.png?raw=true "Run time")