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

| Property       | Example Code                                     |
| -------------- | ------------------------------------------------ |
| Value          | `Cell(1,1).Value = 5`  sets the value of A1 to 5 |
| Fill Color     | `Cell(1,1).Interior.Color = RGB(255, 255, 0)`    |
| Font size      | `Cell(1,1).Font.Size = 14`                       |
| Font name      | `Cell(1,1).Font.Name = "Arial"`                  |
| Formula        | `Cell(1,1).Formula= "=A2*10"`                    |
| Formula (R1C1) | `Cell(1,1).FormulaR1C1="=R2C1*10"`               |
|                |                                                  |



### Worksheet

#### Methods

| Method   | Explanation                  | Example Code                    |
| -------- | ---------------------------- | ------------------------------- |
| Activate | Selects the sheet `"Sheet1"` | `Worksheets("Sheet1").Activate` |
| PrintOut | Prints a sheet               | `ActiveSheet.PrintOut`          |
|          |                              |                                 |
|          |                              |                                 |





