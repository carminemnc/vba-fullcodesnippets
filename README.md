# VBA Snippets

This extension is most comprehensive repository for VBA Snippets.
Starting from basic code snippets (If/Else,For Each,etc.) to entire code snippets useful in your daily work routine.

Originally forked from [spences10's repository](https://github.com/spences10/vba-snippets).


![plot](overview.gif)

## Contacts & Links
You can contact me in the following ways: 
- Mail : [carmine.mnc@gmail.com](mailto:carmine.mnc@gmail.com)
- Github : [carminemnc](https://github.com/carminemnc)
- [Source Code](https://github.com/carminemnc/vba-fullcodesnippets)

## Table of contents

<!-- TOC depthFrom:2 -->

- [Installation](#installation)
- [Overview](#overview)
	- [Basic snippets](#basic-snippets)
	- [Advanced snippets](#advanced-snippets)


<!-- /TOC -->
## Installation
Launch VS Code Quick Open (Ctrl+P), paste the following command, and press enter.
```
ext install vba-fullcodesnippets
```

## Basic Snippets

* Dim declarations
* If/else
* For/While
* Sub/Function with errHandler
* Select Case
* etc.

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
Private Sub func()
'
End Sub
```
#### Private Function code block ```Function```
```vb
Private Function func(input)
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

Some useful entire snippets such as opening a Word file, splitting a worksheet..

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
Dim lr as Long: lr= ActiveWorkbook.ActiveSheet.Cells(Rows.Count,\"A\").End(xlUp).Row
```
#### Last column on the first row ```lc```
```vb
Dim lc as Long: lc= ActiveWorkbook.ActiveSheet.Cells(1,Columns.Count).End(xlToLeft).Column
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

#### Public Function to convert numbers to letter (useful for column) ```#toletter```
```vb
Public Function NumberToLetter(Number As Long) As String
NumberToLetter = Split(Cells(1, Number).Address(True, False),"$")(0)
End Function
```

#### Public Function to convert letter to numbers```letterto#```
```vb
Public Function LetterToNumber(Letter as String) As Long
LetterToNumber = Range(Letter & 1).Column
End Function
```
#### Retrieve Environment Username```user```
```vb
Environ("Username")
```
#### Retrieve ActiveWorkbook Path```curpath```
```vb
ActiveWorkbook.Path
```
#### Create a variant array with column values```cltoarray```
```vb
Dim myArr() as Variant,i as Long
Dim col as String: col="A" ' Define here which column take
Dim start as Long: start=1 ' Here set start=2 if you have headers
Dim lr as Long: lr = ActiveWorkbook.ActiveSheet.Cells(Rows.Count,col).End(xlUp).Row

ReDim myArr(start to lr)

For i=start to lr ' If you have header on your Worksheet set i=2
    myArr(i) = ActiveWorkbook.ActiveSheet.Cells(i,col).Value
Next i

'Debug.Print Join(myArr,",") 'Transposing array
```
#### Open Word Istance ```openword```
```vb
Dim WordApp as Object
Set WordApp = CreateObject("Word.Application")
WordApp.Visible = True
WordApp.Documents.Add
```

#### Open Word File ```openwordfile```
```vb
Dim WordApp as Object,WordDoc as Object
Dim filePath as String: filepath= "C:\Desktop\MyDocument.docx" ' Here insert the filepath
Set WordApp = CreateObject("Word.Application")
WordApp.Visible = True
Set WordDoc = WordApp.Documents.Open(filePath)
```

#### Open URL ```openurl```
```vb
Dim url as String: url= "https://www.google.com"
ActiveWorkbook.FollowHyperlink Address:= url,NewWindow:=True
```

#### Split the current sheet based on column values ```splitsheet```
```vb

Dim lr as Long,lc as String,key as Variant
Dim col as String: col="A" 'Here insert column letter for filtering
Set Unique = CreateObject("Scripting.Dictionary")

With ActiveWorkbook.ActiveSheet

    lr = .Cells(Rows.Count,col).End(xlUp).Row
    lc = NumberToLetter(.Cells(1,Columns.Count).End(xlToLeft).Column)

    Set Data = .Range(col & "1:" & col & lr)

    On Error Resume Next
        For x=2 to lr
            Unique.Add Data(x,1).Value,1
        Next x
    On Error GoTo 0

    For Each key In Unique.Keys
        .Range(col & "1:" & col & lr).AutoFilter, Field:=.Range(col & 1).Column,Criteria1:=key,Operator:=xlFilterValues
        LRFilt = .Range(col & Rows.Count).End(xlUp).Row
        .Range(col & "1:" & lc & LRfilt).SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add(After:=Sheets(ActiveSheet.name)).name = key
        ActiveSheet.Paste
        Cells.EntireColumn.AutoFit
    Next key

End With
```

#### Attach current file to a new mail ```attachthis```
```vb
Application.Dialogs(xlDialogSendMail).Show
```
#### AutoFit Columns ```fitcolumns```
```vb
Cells.Select
Cells.EntireColumn.AutoFit
```
#### AutoFit Rows ```fitrows```
```vb
Cells.Select
Cells.EntireRow.AutoFit
```
#### Copy Active Worksheets into a new Workbook ```copysheet```
```vb
ThisWorkbook.ActiveSheet.Copy Before:=Workbooks.Add.Worksheets(1)
```
#### Refresh all Pivot Tables ```refreshpivots```
```vb
Dim pt As PivotTable

For Each pt In ActiveWorkbook.PivotTables
    pt.RefreshTable
Next pt
```





