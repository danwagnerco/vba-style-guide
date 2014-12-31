This document is largely inspired by the [Ruby Style Guide](https://github.com/bbatsov/ruby-style-guide)

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
