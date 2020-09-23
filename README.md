<div align="center">

## ADO Recordset to Excel


</div>

### Description

Exports an ADO recordset to Microsoft Excel.
 
### More Info
 
ADO Recordset

When done Excel is left open for user interact. Remember to reference Microsoft Excel Object and ActiveX Data Object Libraries in your Project.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Frank Ramos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/frank-ramos.md)
**Level**          |Intermediate
**User Rating**    |4.9 (98 globes from 20 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/frank-ramos-ado-recordset-to-excel__1-37697/archive/master.zip)





### Source Code

```
Public Sub Recordset2Excel(rstSource As ADODB.Recordset)
Dim xlsApp As Excel.Application
Dim xlsWBook As Excel.Workbook
Dim xlsWSheet As Excel.Worksheet
Dim i, j As Integer
 ' Get or Create Excel Object
 On Error Resume Next
 Set xlsApp = GetObject(, "Excel.Application")
 If Err.Number <> 0 Then
  Set xlsApp = New Excel.Application
	Err.Clear
 End If
 ' Create WorkSheet
 Set xlsWBook = xlsApp.Workbooks.Add
 Set xlsWSheet = xlsWBook.ActiveSheet
 ' Export ColumnHeaders
 For j = 0 To rstSource.Fields.Count
  xlsWSheet.Cells(2, j + 1) = rstSource.Fields(j).Name
 Next j
 ' Export Data
 rstSource.MoveFirst
 For i = 1 To rstSource.RecordCount
  For j = 0 To rstSource.Fields.Count
   xlsWSheet.Cells(i + 2, j + 1) = rstSource.Fields(j).Value
  Next j
  rstSource.MoveNext
 Next i
 rstSource.MoveFirst
 ' Autofit column headers
 For i = 1 To rstSource.Fields.Count
  xlsWSheet.Columns(i).AutoFit
 Next i
 ' Move to first cell to unselect
 xlsWSheet.Range("A1").Select
 ' Show Excel
 xlsApp.Visible = True
 Set xlsApp = Nothing
 Set xlsWBook = Nothing
 Set xlsWSheet = Nothing
End Sub
```

