# <span style="color:blue">VBA - Basics</span>

### Table of Contents

1. [Introduction](#1)
2. [Basic Syntax Reference](#2)
	2.1 [Variables](#2.1)
	2.2 [Data Types](#2.2)
	2.3 [Operators](#2.3)
	2.4 [Functions/Subs](#2.4)

## Introduction <a name="1"></a>
	
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

```
Dim SomeNumber As Integer
SomeNumber = 5
```

Keep in mind you cannot assign a value which is of a different type to a declared variable. For example, in the example above trying to do

```
SomeNumber = "Hi"
```
will return an error as "Hi" is of type `String` and not `Integer`

For numeric values, also note the range limits of each type. Using the previous example, if I tried to do:

```
SomeNumber = 100000
```

I would get an error, since the maximum value of an `Integer` is 32,767 , but I am trying to assign a value of 100,000 to it


### 2.2 Data types <a name="2.2"></a>

Here is a list of the different data types supported by VBA! For day to day use you will not need to use most of these so don't worry if it looks complicated!

You will most commonly use these data type so be familiar with them!
`String` `Long` `Boolean` `Single`

#### Numeric

<table style="width:100%">
	<tr>
		<th> Type </th>
		<th> Range</th>
	</tr>
	<tr>
		<td>  Byte </td>
		<td> 0 - 255 </td>
	</tr>
	<tr>
		<td>  Integer </td>
		<td> -32,768 to 32,767 </td>
	</tr>
	<tr>
		<td>  Long </td>
		<td> -2,147,483,648 to 2,147,483,647(integer)  </td>
	</tr>
	<tr>
		<td> LongLong (only on 64-bit systems)</td>
		<td> -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807(integer)</td>
	</tr>
	<tr>
		<td>LongPtr</td>
		<td> Long on 32-bit, LongLong on 64-bit</td>
	</tr>
	<tr>
		<td>  Single </td>
		<td> &#177; -1.401298E-45  to &#177; 3.402823 x 10<sup>38</sup></td>
	</tr>
	<tr>
		<td>  Double </td>
		<td> &#177; 4.94065 x 10<sup>-324</sup> to &#177; 1.79769 x 10<sup>308</sup></td>
	</tr>
	<tr>
		<td>  Currency </td>
		<td> &#177; 922,337,203,685,477.5808  </td>
	</tr>
	<tr>
		<td>  Decimal </td>
		<td> &#177; 79,228,162,514,264,337,593,543,950,335 without decimal or up to 28 decimal places</td>
	</tr>

</table>

#### Non Numeric 

<table style="width:100%">
	<tr>
		<th> Type </th>
		<th> Range</th>
	</tr>
	<tr>
		<td> String </td>
		<td> 0 - 2 billion characters </td>
	</tr>
	<tr>
		<td> Boolean </td>
		<td> True/False </td>
	</tr>
	<tr>
		<td> Date </td>
		<td> January 1, 100 to December 31, 9999 </td>
	</tr>
	<tr>
		<td> Object </td>
		<td> embedded objects </td>
	</tr>
	<tr>
		<td> Variant </td>
		<td> up to Double or String </td>
	</tr>
</table>


### 2.3 Operators <a name="2.3"></a>

#### Arithmetic

<table style="width:100%">
	<tr>
		<th>Operator</th>
		<th>Description</th>
		<th>Examples(where necessary)</th>
	</tr>
	<tr>
		<td> + </td>
		<td> Addition </td>
		<td></td>
	</tr>
	<tr>
		<td> - </td>
		<td> Subtraction </td>
		<td></td>
	</tr>
	<tr>
		<td> * </td>
		<td> Multiplication</td>
		<td></td>
	</tr>
	<tr>
		<td> / </td>
		<td> Division </td>
		<td></td>
	</tr>
	<tr>
		<td> \ </td>
		<td> Floor division </td>
		<td> 5 \ 2 = 2</td>
	</tr>
	<tr>
		<td> % </td>
		<td> Modulo (Remainder function) </td>
		<td> 5 & 2 = 1 </td>
	</tr>
	<tr>
		<td> ^ </td>
		<td> Exponentiation </td>
		<td></td>
	</tr>
	
</table>

#### Comparison

<table style="width:100%">
	<tr>
		<th>Operator</th>
		<th>Description</th>
	</tr>
	<tr>
		<td> = </td>
		<td> Equal to </td>
	</tr>
	<tr>
		<td> &lt; &gt; </td>
		<td> Not equal to (!= in other languages) </td>
	</tr>
	<tr>
		<td> > / >= </td>
		<td> Greater than / Greater than or equal to </td>
	</tr>
	<tr>
		<td> &lt; / &lt;= </td>
		<td> Less than / Less than or equal to </td>
	</tr>	
	
</table>


#### Logical

<table style="width:100%">
	<tr>
		<th>Operator</th>
		<th>Description</th>
	</tr>
	<tr>
		<td> AND </td>
		<td> Returns True if all statements are true </td>
	</tr>
	<tr>
		<td> OR </td>
		<td> Returns True if <b>ANY</b> statement is true </td>
	</tr>
	<tr>
		<td> NOT </td>
		<td> Returns True if no statements are true </td>
	</tr>
	<tr>
		<td> XOR </td>
		<td> Returns True if <b>ONE</b> statement is true </td>
	</tr>	
	
</table>

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

```
1. Function Distance(v as Double, t as Double) as Double
2.
3. Distance = v * t 
4.
5. End Function
```

As you can see, there are 3 parts to a function. In Line 1, you give the function a name, `Distance` and give it the arguments required to calculate the value, `v` and `t`

In Line 3, you define the code required to calculate the distance, `v * t`

Finally, you end the function with `End Function` as per Line 5

#### Calling a function/sub

Now that you have your brand new, shiny function, all you have to do is to call it in your code!

```
1. Dim s as Double
2. 
3. s = Distance(3 , 5)
```

The above code sets the value of `s` to the return value of `Distance` with input arguments `v = 3` and `t = 5`. Therefore the expected value of `s` would be 15

