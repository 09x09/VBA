# <span style="color:blue">VBA - Basics</span>

Author : 09x09

Learning how to program is always easiest when you have something to practice with! Don't just read this like a textbook cos it's not going to end well~ (Also because this guide does not include every function out there, only those I think would be most useful for day to day use)

----

### Table of Contents

1. [Introduction](#1)

2. [Basic Syntax Reference](#2)

    2.1 [Variables](#2.1) 

    2.2 [Data Types](#2.2) 

    2.3 [Operators](#2.3)

    2.4 [Functions/Subs](#2.4) 

    2.5 [Loops](#2.5) 

    2.6 [Conditionals](#2.6)

    2.7 [Arrays](#2.7)

    2.8 [Strings](#2.8)

    2.9 [Objects](#2.9)

    2.10 [Numbers](#2.10)

    2.11 [Date and Time](#2.11)

    

## Introduction <a name="1"></a>

VBA is a really powerful tool in Excel. It allows you to automate long tedious tasks like moving data and fixing poor data entry practices *cough cough* as well as create user interaction with your spreadsheets. With how prevalent Excel use is it is definitely a worthwhile language to pick up!



## Basic Syntax Reference <a name="2"></a>

### 2.1 Variables <a name="2.1"></a>
#### Declaring Variables

Because VBA is a pretty old language, the type of each variable has to be declared along with it's name, like so:

```
Dim <name> as <type>
```

It's good practice to use meaningful variable names. For example, `a` is a poor variable name, while something like `car_speed` is much better as it tells you what the variable is used for

#### Assigning variables

After you have declared a variable, you can assign a value to it!

```vb
Dim SomeNumber As Integer
SomeNumber = 5
```

Keep in mind you cannot assign a value which is of a different type to a declared variable. For example, in the example above trying to do

```vb
SomeNumber = "Hi"
```
will return an error as "Hi" is of type `String` and not `Integer`

For numeric values, also note the range limits of each type. Using the previous example, if I tried to do:

```vb
SomeNumber = 100000
```

I would get an error, since the maximum value of an `Integer` is 32,767 , but I am trying to assign a value of 100,000 to it


### 2.2 Data types <a name="2.2"></a>

Here is a list of the different data types supported by VBA! For day to day use you will not need to use most of these so don't worry if it looks complicated!

You will most commonly use these data type so be familiar with them!
`String` `Long` `Boolean` `Double`

#### Numeric

| Type                  | Range                                                        | Conversion |
| --------------------- | ------------------------------------------------------------ | ---------- |
| Byte                  | 0-255                                                        | `CByte`    |
| Integer               | -32,768 to 32,767                                            | `CInt`     |
| Long                  | -2,147,483,648 to 2,147,483,647(integer)                     | `CLng`     |
| LongLong(only 64-bit) | -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807(integer) |            |
| LongPtr               | Long on 32-bit, LongLong on 64-bit                           |            |
| Single                | &#177; -1.401298E-45  to &#177; 3.402823 x 10<sup>38         | `CSng`     |
| Double                | &#177; 4.94065 x 10<sup>-324</sup> to &#177; 1.79769 x 10<sup>308 | `CDbl`     |
| Currency              | &#177; 922,337,203,685,477.5808                              |            |
| Decimal               | &#177; 79,228,162,514,264,337,593,543,950,335 without decimal or up to 28 decimal places | `CDec`     |

#### Non Numeric 

| Type    | Range                               | Conversion |
| ------- | ----------------------------------- | ---------- |
| String  | 0 - 2billion characters             | `CStr`     |
| Boolean | True/False                          | `CBool`    |
| Date    | January 1, 100 to December 31, 9999 | `CDate`    |
| Object  | embedded objects                    | `CObj`     |
| Variant | up to Double or String              |            |



### 2.3 Operators <a name="2.3"></a>

#### Arithmetic

| Operator | Description      | Examples(where necessary) |
| -------- | ---------------- | ------------------------- |
| +        | Addition         |                           |
| -        | Subtraction      |                           |
| *        | Multiplication   |                           |
| /        | Division         |                           |
| \        | Integer Division | 5 \ 2 = 2                 |
| %        | Modulo           | 5 % 2 = 1                 |
| ^        | Exponentiation   |                           |

#### Comparison

| Operator | Description                             |
| -------- | --------------------------------------- |
| =        | Equal to                                |
| <>       | Not Equal To                            |
| > / >=   | Greater than / Greater than or equal to |
| < / <=   | Less than / Less than or equal to       |



#### Logical

| Operator | Description                                      |
| -------- | ------------------------------------------------ |
| AND      | Returns `True` only if ALL statements are `True` |
| OR       | Returns `True` if one statement is `True`        |
| NOT      | Returns `True` only if NO statements are `True`  |
| XOR      | Returns `True` only if ONE statements is `True`  |



#### Concatenation



<table style="width:100%">
	<tr>
		<th>Operator</th>
		<th>Examples</th>
	</tr>
	<tr>
		<td> + </td>
		<td> Avoid using if possible as it uses the same symbol as the addition operator </td>
	</tr>
	<tr>
		<td rowspan=4> & </td>
		<td> 1 & 2 = "12" </td>
	</tr>
	<tr>
		<td> "1" & 2 = "12"</td>
	</tr>
	<tr> 
		<td>1 & "Hello" = "1Hello"</td>
	</tr>
	<tr>
		<td> "Hello" & "World" = "HelloWorld" </td>
	</tr>
	
</table>

### 2.4 Functions and Subs <a name="2.4"></a>

Functions and subs(subroutines) make life easier by allowing you to reuse lines of code in different places instead of having to type out the full code each time

Functions are different from subs in that functions have a return value, while subs do not. This means that you can feed the output value of a function into another function

Subs are used when you want Excel to do something

#### Declaring a function/sub

Let's do a simple example, say I want to calculate the distance travelled by a car

```vb
1. Function Distance(v as Double, t as Double) as Double
2.
3. 	Distance = v * t 
4.
5. End Function
```

As you can see, there are 3 parts to a function. In Line 1, you give the function a name, `Distance` and give it the arguments required to calculate the value, `v` and `t`

In Line 3, you define the code required to calculate the distance, `v * t`. The indentation here is not necessary but it is good practice for code readability.

Finally, you end the function with `End Function` as per Line 5

Declaring a sub is similar to a function, as so:

```vb
1. Sub Distance(v as Double, t as Double)
2.
3. 	MsgBox = v * t 
4.
5. End Sub
```

*the MsgBox function produces a message box with a value in it

#### Calling a function

Now that you have your brand new, shiny function, all you have to do is to call it in your code!

```vb
1. Dim s as Double
2. 
3. s = Distance(2 , 5) + 5
4. MsgBox s
```

The above code sets the value of `s` to the return value of `Distance` with input arguments `v = 2` and `t = 5` added to 5.



#### Calling a sub

Place a command button on your worksheet and add the following lines to it

```vb
1. Distance 2 , 5
```

This should produce a message box with a value of 10 in it



### 2.5 Loops <a name="2.5"></a>

Loops are used  when you want to repeat a task for a fixed number of times or until some condition is fulfilled

#### For Loops

For loops will perform a task for a set number of times. This is the structure of a for loop:

```vb
1.Dim i as Integer
2.
3.For i = 1 to 10
4.	<some code> 
5.Next i
```



#### While Loops

While loops will perform a task while some condition remains `True` . This is the structure of a while loop:

```vb
1.Dim i as Integer
2.i = 1
3.
4.Do While i < 10
5.	<some code>
6.	i = i + 1
7.Loop
```

Here is another version of the while loop:

```vb
1.Dim x as Boolean
2.x = True
3.
4.Do While x = True
5.	<some code>
6.	If [conditional] Then
7.        x = False
8.  End If
9.Loop    
```

Note that the line right before `Loop` acts as the control for the loop condition. If this part is missing the loop becomes an infinite loop

#### Nested Loops

Loops can also be nested, as shown below:

```vb
1.Dim i, j as Integer
2.i, j = 1
3.
4.Do While i < 10
5.	For j = 1 to 10    
6.		<some code>
7.	Next j
8.	i = i + 1
9.Loop
```



### 2.6 Conditionals <a name="2.6"></a>

Conditionals, as the name suggests, allows you to run a piece of code only when certain conditions are fulfilled. This is the basic structure of a conditional:

```vb
1.If [condition1] Then
2.   
3.   <your code here>
4.    
5.ElseIf [condition2] Then
6.   
7.    <more code>
8.    
9.Else
10.    
11.    <even more code>
12.    
13.End If
```

`ElseIf` statements are completely optional. Furthermore, there is no limit to the number of `ElseIf` statements you can have in a conditional



### 2.7 Arrays<a name="2.7"></a>

Arrays are collections of objects. Arrays can be static or dynamic, and can also be nested (multi dimensional)

#### Static Arrays

This is how you declare a static array as well as store values in it!

```vb
1.Dim Cars(1 to 5) as String
2.
3.Cars(1) = "Honda"
4.Cars(2) = "Toyota"
5.Cars(3) = "McLaren"
etc
```



#### Dynamic Arrays

Dynamic arrays work like static arrays, except that their size is not fixed. This is how you declare a dynamic array!

```vb
1.Dim Cars() as String
```

To resize a dynamic array we use the `ReDim` function

```vb
1.ReDim Cars(4)
```



### 2.8 Strings<a name="2.8"></a>

Strings are expressions enclosed within quotation marks `"`. Strings can include numbers as well as symbols. So `"1"` is a string, but `1` is not.

#### String functions

#### Get length of string

| Purpose | Function | Example        | Expected Value |
| ------- | -------- | -------------- | -------------- |
| Length  | Len()    | `Len("Apple")` | 5              |

#### Reversing a string

| Purpose | Function     | Example               | Expected Value |
| ------- | ------------ | --------------------- | -------------- |
| Reverse | StrReverse() | `StrReverse("Apple")` | "elppA"        |

#### Extract values from strings

| Purpose             | Function | Example              | Expected Value |
| ------------------- | -------- | -------------------- | -------------- |
| Extract from left   | Left()   | `Left("Apple",3)`    | "App"          |
| Extract from right  | RIght()  | `Right("Apple",3)`   | "ple"          |
| Extract from middle | Mid()    | `Mid("Apple", 2, 2)` | "ppl"          |

#### Case conversion

| Purpose                | Function  | Example                          | Expected output |
| ---------------------- | --------- | -------------------------------- | --------------- |
| Convert to upper case  | UCase()   | `UCase("apple")`                 | "APPLE"         |
| Convert to upper case  | StrConv() | `StrConv("apple", vbUpperCase)`  | "APPLE"         |
| Convert to lower case  | LCase()   | `LCase("APPLE")`                 | "apple"         |
| Convert to lower case  | StrConv() | `StrConv("APPLE", vbLowerCase)`  | "apple"         |
| Convert to proper case | StrConv() | `StrConv("apple", vbProperCase)` | "Apple"         |

#### Removing blank spaces

| Purpose | Function | Example           | Expected Value |
| ------- | -------- | ----------------- | -------------- |
| Left    | LTrim()  | `LTrim(" Apple")` | "Apple"        |
| Right   | RTrim()  | `RTrim("Apple ")` | "Apple"        |
| Both    | Trim()   | `Trim(" Apple ")` | "Apple"        |

#### Finding a substring

Finding a substring within a string can be done with InStr() and InStrRev(). The only difference between the two is that InStrRev() searches from the end of the string. Both functions return the position of the first match found.

| Function                              | Example                     | Expected Value |
| ------------------------------------- | --------------------------- | -------------- |
| InStr([optional], String 1, String 2) | InStr(4, "John Smith", "h") | 10             |
| InStr(String 1, String 2)             | InStr("John Smith", "h")    | 3              |



### 2.9 Objects <a name="2.9"></a>

An object is an identifier for a collection of methods and properties. Properties describe the object, while methods are functions which belong to the object. Referring to these methods and properties is done through the use of `.`, as in `Object.property` or `Object.method`

For example, let's say we have a `Car` object. A property of this `Car`object might be it's `color` . So `Car.color` will refer to the color of our `Car`

A method of our `Car` might be the function `Distance(v,t)` which returns the distance travelled by the car. So if we want to call this method for a car travelling at `v = 10` for `t = 1`  we will call `Car.Distance(10, 1)`

Objects can be nested within other objects as well!

#### Objects in Excel

Almost everything in excel is an object, which is why this subsection is here



### 2.10 Numbers<a name="2.10"></a>

#### Floats

A float is the decimal representation of a number. Integers can also be expressed as floats, for example, `1` is an integer while `1.0` is a float

#### Floating point arithmetic

Because memory is finite, and binary representations of some floats may not be, often times there will be a rounding error involved in storing the float.  As a result, you may not get expected results when performing some operations.

For example, if you try to compare floats directly, often times you will not get the expected output. 

```vb
If 0.5-0.4-0.1 = 0 Then
    MsgBox("True")
    
Else
    MsgBox("False")
    
End If
```

Pasting the above into a command button will yield a message stating "False" instead of the expected "True"

To get around this, modify the comparison to the one below:

```vb
If Abs(0.5-0.4-0.1 - 0) < 10^-6 Then

```

This will allow you to get expected results

For more information on floating point arithmetic in excel 

https://www.microsoft.com/en-us/microsoft-365/blog/2008/04/10/understanding-floating-point-precision-aka-why-does-excel-give-me-seemingly-wrong-answers/



### Date and Time<a name="2.11"></a>

Date and time in VBA can be expressed as a number or in conventional date and time formats we are familiar with

#### Initialising a date object

To initialise a date object we use the DateValue function

```vb
1. Dim someDate As Date
2.
3.someDate = DateValue("1 Jan 2018")
```

To initialise a time object we use the TimeValue function instead

Internally, Excel stores dates and times as a number, which is the number of days since January 0, 1900. For example, 40 745 is the date 21/7/2011

Manipulating dates and times can be useful, so here are the methods which are available!

| Purpose                                        | Function                  | Example                      | Result               |
| ---------------------------------------------- | ------------------------- | ---------------------------- | -------------------- |
| Get year                                       | Year()                    | `Year(someDate)`             | 2018                 |
| Get month                                      | Month()                   | `Month(someDate)`            | 1                    |
| Get day                                        | Day()                     | `Day(someDate)`              | 1                    |
| Add days to a date                             | DateAdd()                 | `DateAdd("d", 7, someDate)`  | 8/1/2018             |
| Add months to a date                           | DateAdd()                 | `DateAdd("m", 7, someDate)`  | 1/8/2018             |
| Add years to a date                            | DateAdd()                 | `DateAdd("y", 7, someDate)`  | 1/1/2025             |
| Get current date and time                      | Now                       | `Msgbox Now`                 | 8/8/2018 10:48:15 AM |
| Get Hour/Minutes/Seconds                       | Hour()/Minute()/Seconds() | -                            | -                    |
| Difference between 2 dates (days/months/years) | DateDiff()                | `DateDiff("d",Date1, Date2)` | -                    |





