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
      'do something'
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

* Prepend all variables with a 3-letter code to indicate its type. This is commonly referred to as Hungarian Notation (or, more accurately, [_Apps Hungarian_](http://en.wikipedia.org/wiki/Hungarian_notation)

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

![Function Input Variable](http://github/danwagnerco/vba-style-guide/raw/master/images/function-input-name.png)

* When working with integers, use `Long` instead of `Integer`.

  ```vb
  'Bad
  Dim intValue As Integer

  'Good
  Dim lngValue As Long
  ```
