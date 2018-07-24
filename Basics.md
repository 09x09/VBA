# <span style="color:blue">VBA - Basics</span>

### Data types

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
		<td> &#177; 2,147,483,648  </td>
	</tr>
	<tr>
		<td>  Single </td>
		<td> &#177;  3.402823 x 10<sup>38</sup></td>
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

### Variables
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


### Operators

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