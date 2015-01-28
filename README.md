This document is largely inspired by the [Ruby Style Guide](https://github.com/bbatsov/ruby-style-guide).

# Prelude
> No one man should have all that power.<br/>
> -- Kanye West

This is an evolving document. Submit a pull request and start the conversation!

# VBA Style Guide

## Source Code Layout

> Not complicated, it's simple.<br/>
> -- Big Sean

* Limit lines to 80 characters.

* Use 4-character tabs to indent.

  ```vb
  'Bad
  If blnSomething Then
    Msgbox "True" '<~ 2-character indents
  Else
    Msgbox "False"
  End If

  'Good
  If blnSomething Then
      Msgbox "True" '<~ 4-character indents
  Else
      Msgbox "False"
  End If
  ```

* All conditionals, loops and blocks should be indented.

  ```vb
  'Bad
  If lngNumber >= 0 Then
  Msgbox "Yep"
  Else
  Msgbox "Nope"
  End

  'Good
  If lngNumber >= 0 Then
      Msgbox "Yep"
  Else
      Msgbox "Nope"
  End

  'Bad
  For lngIndex = 1 To lngLastRow
  lngCounter = lngCounter + lngIndex
  Next lngIndex

  'Good
  For lngIndex = 1 To lngLastRow
      lngCounter = lngCounter + lngIndex
  Next lngIndex

  'Bad
  With wksSource
  Set rngSource = .Range(.Cells(1, 1), .Cells(lngLastRow, 1))
  End With

  'Good
  With wksSource
      Set rngSource = .Range(.Cells(1, 1), .Cells(lngLastRow, 1))
  End With
  ```

## Variables and Naming

> Oh that looks like what's-her-name, chances are it's what's-her-name.<br/>
> -- Drake

* Use `Option Explicit` to mandate variable declaration.

  ```vb
  'Bad
  Public Sub MyMacro()
      'do something
  End Sub

  'Good
  Option Explicit
  Public Sub MyMacro()
      'do something
  End
  ```

* Declare variable types explicitly.

  ```vb
  'Bad
  Dim MyNumber
  Dim MyBlock
  Dim MyVariable

  'Good
  Dim MyNumber As Long
  Dim MyBlock As Range
  Dim MyVariable As Variant
  ```

* Prepend all variables with a 3-letter code to indicate its type. This is commonly referred to as Hungarian Notation (or, more accurately, [_Apps Hungarian_](http://en.wikipedia.org/wiki/Hungarian_notation))

  | Variable Type      | 3-Letter Code |
  | ------------------ | ------------- |
  | Workbook           | wbk           |
  | Worksheet          | wks           |
  | Long               | lng           |
  | Double             | dbl           |
  | String             | str           |
  | Range              | rng           |
  | Boolean            | bln           |
  | Object             | obj           |
  | FileDialog         | fdo           |
  | Collection         | col           |
  | Variant            | var           |
  | Comment            | cmt           |
  | ChartObject        | cho           |
  | Shape              | shp           |

  **EXCEPTION:** input variables to function should NOT have a 3-letter code. These variable types can be identified trivially by Intellisense and should be named to maximize readability:

  ![Image](images/function-input-name.png)

* When working with whole numbers, use `Long` instead of `Integer`.

  ```vb
  'Bad
  Dim intValue As Integer

  'Good
  Dim lngValue As Long
  
  'Integers are 16-bit and can only store values up to 32,767
  lngValue = 50000 '<~ no issue
  intValue = 50000 '<~ overflow error
  ```
  
  Under the covers, your CPU converts `Integer`-type variables into `Long`-type variables, does the math then converts the resulting `Long` back to an `Integer`. Avoid the debugging headache and default all integers to `Long`.

* Name **local** variables in `CamelCase`. (Keep acronyms like HTTP, RFC and XML uppercase.)

  ```vb
  'Bad
  Dim str_my_variable As String
  Dim strmyvariable As String
  Dim strHttp As String

  'Good
  Dim strMyVariable As String
  Dim strHTTP As String
  ```
  
* Name **global** variables in `SCREAMING_SNAKE_CASE`. (Keep acronyms like HTTP, RFC and XML uppercase.)

  ```vb
  'Bad
  Dim str_error_message As String
  Dim strErrorMessage As String
  Dim lngCONSTANT As Long
  Dim lngHttpAcceptedCode As Long
   
  'Good
  Dim str_ERROR_MESSAGE As String
  Dim lng_CONSTANT As Long
  Dim lng_HTTP_ACCEPTED_CODE As Long
  ```

* Generic 3-letter variable names are OK for common iterators.

  ```vb
  'OK
  Dim wks As Worksheet
  
  For Each wks In ThisWorkbook.Worksheets
      'do stuff to each sheet
      Msgbox (wks.Name)
  Next wks
  ```

**NEED TO FLESH THESE OUT MORE**
* All functions should be stored in a unique module
* All functional subroutines should be stored in a unique module
* All public subroutines should be stored in a unique module
* All global constants should be stored in a unique module

## Syntax

> I'mma be fresh as hell if the Feds watching.<br/>
> -- 2 Chainz & Pharrell

* Close `For...Next` loops with the iterative variable after `Next`.

  ```vb
  'Bad
  For Each wks in ThisWorkbook.Worksheets
      'do stuff with each worksheet
  Next
  
  'Good
  For Each wks in ThisWorkbook.Worksheets
      'do stuff with each worksheet
  Next wks
  ```
* 
