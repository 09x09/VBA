# Excel VBA

Author: 09x09



---

## Pre requisites

1. Being familiar with the basics of VBA
2.  Being at least somewhat familiar with Excel



## Table of Contents

1. Terminology

2. Objects

3. [~~Basic~~ Syntax Reference](#3)

   3.1[Workspace](#3.1)

   3.2[Worksheet](#3.2)

   3.3[Cells](#3.3)

   3.4[UserForms](#3.4)

## 1. Terminology <a name="1"></a>

#### R1C1 

R1C1 is the other method of referencing cells, instead of using the A1 notation. In R1C1, each cell is referenced using its row and column number. So cell `A1` is `R1C1` , `C3` is `R3C3` etc.

Relative referencing in R1C1 is represented by square brackets `[]` around the numbers. For example, if our current active cell is `R3C3` (`C3`), then `R[-1]C[1]` is the cell `R2C4` , which is `D2`



## 2. Objects <a name="2"></a>

#### What is an object?

An object is an identifier for a collection of methods and properties. Properties describe the object, while methods are functions which belong to the object. Referring to these methods and properties is done through the use of `.`, as in `Object.property` or `Object.method`

For example, let's say we have a `Car` object. A property of this `Car`object might be it's `color` . So `Car.color` will refer to the color of our `Car`

A method of our `Car` might be the function `Distance(v,t)` which returns the distance travelled by the car. So if we want to call this method for a car travelling at `v = 10` for `t = 1`  we will call `Car.Distance(10, 1)`



## 3. ~~Basic~~ Syntax Reference<a name="3"></a>

### 3.1 Workbook<a name="3.1"></a>

### 3.2 Worksheet<a name="3.2"></a>

### 3.3 Cells<a name="2.3"></a>

#### Cell Referencing





`



### 3.4 UserForms<a name="3.4"></a>

