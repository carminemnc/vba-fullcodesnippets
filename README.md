# VBA Snippets

Originally forked from [spences10's repository](https://github.com/spences10/vba-snippets).

This extension contains almost all useful snippets that you can use in VBA.


<!-- TOC depthFrom:2 -->

- [Overview](#overview)
	- [Basic snippets](#basic-snippets)
	- [Advanced snippets](#advanced-snippets)
- [Installation](#installation)
- [Contacts](#contacts)
- [Links](#links)

<!-- /TOC -->

## Overview
Below you will find all the snippets presents in this extension.

For each extension in ```"code markdown"```, the snippet shortcut in Visual Studio Code.


## Basic Snippets

* Dim declarations
* If/else
* For/While
* Sub/Function with errHandler
* Select Case

#### Dim Declarations ```Dim```
```vb
Dim arr()
Dim bol As Boolean
Dim lng As Long
Dim dbl As Double
Dim str As String
Dim obj As Object
Private
```
#### If ```If```
```vb
If condition Then

End If
```
#### Else ```Else```
```vb
If condition Then

Else

End If
```
#### ElseIf ```ElseIf```
```vb
ElseIf condition2 Then
```
#### With code block ```With```
```vb
With

End With
```
#### For Next Loop code block ```for```
```vb
For i = lower To upper

Next i
```
#### For Each Loop code block ```ForEach```
```vb
For Each variable In collection

Next variable
```
#### Do Loop While code block ```DoLoopWhile```
```vb
Do

Loop While condition
```
#### DoWhile code block ```DoWhile```
```vb
Do While condition

Loop
```
#### While Wend code block ```While```
```vb
While condition

Wend
```
#### Sub code block ```Sub```
```vb
Private Sub func(input)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'



End Sub
```
#### Private Function code block ```Function```
```vb
Private Function func(input)
'
' descrip.
'
' @since 1.0.0
' @param {type} [name] descrip.
' @return {type} [name] descrip.
' @see dependencies
'




End Function
```
#### SelectCase code block ```SelectCase```
```vb
Select Case test

  Case lists

    statements

  Case Else

    elseStatement

End Select
```

#### Short Snippets
```vb
UBound
LBound
To
Fix
Int
ReDim
Set
Call
Split
Preserve
Option Explicit
On Error Resume Next
ClearContents
Clear
Columns
Rows
CreateObject
IsEmpty
```


## Advanced Snippets
* Enabling/Disabling Screen Updating
* Mail Code Block
* Last row / Last Column
* ColumnLetter

#### Enabling Screen Updating ```screenon```
```vb
With Application
   .ScreenUpdating = True
End With
```
#### Disabling Screen Updating ```screenoff```
```vb
With Application
   .ScreenUpdating = False
End With
```

#### Last row on the first column ```lr```
```vb
.Cells(Rows.Count,1).End(xlUp).Row
```
#### Last column on the first row ```lc```
```vb
.Cells(1,Columns.Count).End(xlToLeft).Column
```

#### Mail code block ```mail```
```vb
Dim OutApp as Object,OutMail as Object
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = ""
        .Subject = ""
        .CC = ""
        .Body = ""
        .Attachments.Add 'Here paste the filepath
        .Display '.Send to send the mail
    End With

Set OutApp = Nothing
Set OutMail = Nothing
```

#### Public Function to convert numbers to letter (useful for column) ```cltoletter```
```vb
Public Function ColumnLetter(ColumnNumber As Long) As String
ColumnLetter = Split(Cells(1, ColumnNumber).Address(True, False),"$")(0)
End Function
```

## Installation
Launch VS Code Quick Open (Ctrl+P), paste the following command, and press enter.
```
ext install vba-fullcodesnippets
```
## Contacts
You can contact me in the following ways: 
- Mail : [spences10apps@gmail.com](mailto:carmine.mnc@gmail.com)
- Github : [spences10](https://github.com/carminemnc)

## Links
- [Source Code](https://github.com/carminemnc/vba-fullcodesnippets)
- [VS Market](https://marketplace.visualstudio.com/items/spences10.vba-snippets)

