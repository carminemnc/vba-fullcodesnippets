# VBA Snippets

This extension is the most comprehensive repository for VBA Snippets.
Starting from basic code snippets (If/Else,For Each,etc.) to entire code snippets useful in your daily work routine.

Built from a VBA user for VBA users.

![plot](overview.gif)

## Contacts & Links
You can contact me in the following ways: 
- Mail : [carmine.mnc@gmail.com](mailto:carmine.mnc@gmail.com)
- Github : [carminemnc](https://github.com/carminemnc)
- [Source Code](https://github.com/carminemnc/vba-fullcodesnippets)

## Table of contents

<!-- TOC depthFrom:2 -->

- [Installation](#installation)
- [Contents](#contents)
	- [Basic snippets](#basic-snippets)
	- [Advanced snippets](#advanced-snippets)


<!-- /TOC -->
## Installation
Launch VS Code Quick Open (Ctrl+P), paste the following command, and press enter.
```
ext install vba-fullcodesnippets
```

### Basic Snippets


| abbreviation  | description                         |
| :------------ | :---------------------------------- |
| `Dim`         | Dim declarations(Long,Boolean etc.) |
| `If`          | If statement                        |
| `Else`        | Else statement                      |
| `Elseif`      | Else if statement                   |
| `With`        | With code block                     |
| `for`         | For next loop                       |
| `ForEach`     | For each loop code block            |
| `DoWhile`     | Do while loop code block            |
| `Function`    | Function code block                 |
| `SelectCase`  | Select case code block              |
| `cm`          | Comment block                       |
| `Split`       | Split function                      |
| `Sub`         | Sub code block                      |
| `While`       | While Wend code block               |
| `DoLoopWhile` | Do Loop While code block            |
| `Range`       | Range object                        |

### Advanced Snippets

| abbreviation    | description                                     |
| :-------------- | :---------------------------------------------- |
| `screenon`      | Enabling screen updating                        |
| `screenoff`     | Disabling screen updating                       |
| `lr`            | Last row of the first row of the worksheet      |
| `lc`            | Last column of the first row of the worksheet   |
| `mail`          | Create a snippet for building up an email       |
| `#toletter`     | Public function to convert number into letter   |
| `letterto#`     | Public function to convert letter to number     |
| `worksheetloop` | Loop trough current worksheets                  |
| `user`          | Current  user                                   |
| `curpath`       | Active workbook path                            |
| `cltoarray`     | Create a variant array with column values       |
| `openword`      | Open a new word istance                         |
| `openwordfile`  | Open a word file                                |
| `openurl`       | Open a URL                                      |
| `splitsheet`    | Split the current sheet based on column values  |
| `attachthis`    | Attach current workbook to a new mail           |
| `fitcolumns`    | Fit columns width of the worksheet              |
| `fitrows`       | Fit rows width for the worksheet                |
| `copysheet`     | Copy current worksheet into a new workbook      |
| `refreshpivots` | Refresh all pivot tables in the active workbook |

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

#### Public function to convert numbers to letter (useful for column) ```#toletter```
```vb
Public Function NumberToLetter(Number As Long) As String
NumberToLetter = Split(Cells(1, Number).Address(True, False),"$")(0)
End Function
```

#### Public function to convert letter to numbers ```letterto#```
```vb
Public Function LetterToNumber(Letter as String) As Long
LetterToNumber = Range(Letter & 1).Column
End Function
```
#### Loop trough current worksheets ```worksheetloop```
```vb
Dim ws as Worksheet

For Each ws in ActiveWorkbook.Sheets
    Debug.Print ws.Name
Next ws
```

#### Retrieve environment username  ```user```
```vb
Environ("Username")
```
#### Retrieve ActiveWorkbook Path ```curpath```
```vb
ActiveWorkbook.Path
```
#### Create a variant array with column values ```cltoarray```
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
#### Open word istance ```openword```
```vb
Dim WordApp as Object
Set WordApp = CreateObject("Word.Application")
WordApp.Visible = True
WordApp.Documents.Add
```

#### Open word file ```openwordfile```
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
#### Copy active worksheet in a new workbook ```copysheet```
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





