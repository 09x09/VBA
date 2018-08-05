# Excel VBA

Author: 09x09



---

## Pre requisites

1. Being familiar with the basics of VBA
2.  Being at least somewhat familiar with Excel



## Table of Contents

1. Terminology

2. [~~Basic~~ Syntax Reference](#2)

   

## 1. Terminology <a name="1"></a>

#### R1C1 

R1C1 is the other method of referencing cells, instead of using the A1 notation. In R1C1, each cell is referenced using its row and column number. So cell `A1` is `R1C1` , `C3` is `R3C3` etc.

Relative referencing in R1C1 is represented by square brackets `[]` around the numbers. For example, if our current active cell is `R3C3` (`C3`), then `R[-1]C[1]` is the cell `R2C4` , which is `D2`



## 2. ~~Basic~~ Syntax Reference<a name="2"></a>

### Cells

#### Referencing

|            |                    | Example Code     |
| ---------- | ------------------ | ---------------- |
| Many Cells | Range()            | `Range("A1:D4")` |
| One Cell   | Cells(Row, Column) | `Cells(1,1)`     |



#### Properties

| Property       | Example Code                                  |
| -------------- | --------------------------------------------- |
| Value          | `Cell(1,1).Value = 5`                         |
| Fill Color     | `Cell(1,1).Interior.Color = RGB(255, 255, 0)` |
| Font size      | `Cell(1,1).Font.Size = 14`                    |
| Font name      | `Cell(1,1).Font.Name = "Arial"`               |
| Formula        | `Cell(1,1).Formula= "=A2*10"`                 |
| Formula (R1C1) | `Cell(1,1).FormulaR1C1="=R2C1*10"`            |
|                |                                               |



### Range

Range objects are possibly the most widely encountered and used objects when coding in VBA.  The methods and properties of the range object includes those mentioned above. 

Below are some of the commonly used properties and methods. The full list can be found here: https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/range-object-excel

#### Methods

|Method|Explanation|Example|
||||
||



#### Properties

| Properties | Example |
| ---------- | ------- |
|            |         |







### Worksheet

Below is a list of useful properties and methods of the Worksheet object. A full list can be found at https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheet-object-excel

#### Methods

| Method      | Explanation                                | Example Code                                                 |
| ----------- | ------------------------------------------ | ------------------------------------------------------------ |
| Activate    | Selects the sheet `"Sheet1"`               | `Worksheets("Sheet1").Activate`                              |
| PrintOut    | Prints a sheet                             | `ActiveSheet.PrintOut`                                       |
| PivotTables | Returns a PivotTable object on a worksheet | `ActiveSheet.PivotTables("PivotTable1"). _ <br />  PivotFields("Sum of 1994").Function = xlSum ` |
|             |                                            |                                                              |



#### Using Excel Functions in VBA

To use excel functions, the `WorksheetFunction` object is called along with the excel function desired. For example, `WorksheetFunction.Min(Range)` gets the minimum value in a range

Here are some useful excel functions, a full list can be found at https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheetfunction-object-excel 

| Function   | Use                                                  |
| ---------- | ---------------------------------------------------- |
| CountA     | Count number of cells containing data within a range |
| CountBlank | Count empty cells within a range                     |
| IsNumber   | Returns true if cell contains a number               |
|            |                                                      |
|            |                                                      |
|            |                                                      |





