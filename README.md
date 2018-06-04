# ExcelApiPerformanceTests
Simple Project to compare the performance of excel calls via [MS Interop Late Binding](https://msdn.microsoft.com/library/microsoft.office.interop.excel.aspx) and [NetOfficeFw](https://github.com/NetOfficeFw/NetOffice)

The programm starts two threads. One thead will open an excel document via late binding with Microsoft Interop and in the other will use the NetOffice Library.

In order to see the performance it will read around 100x100 cells with each value and formular and write these to a new excel document.

3 ways of reading cells are implemented in order to find a fast solution of accessing cells. Each algrothim writes to its own sheet.
