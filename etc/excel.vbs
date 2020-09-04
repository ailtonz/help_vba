dim app
set app = createobject("Excel.Application")

app.Visible = true

dim wb
set wb = app.workbooks.open("C:\temp\CalculoPreco.xls")

For Each ws In app.Worksheets
	If ws.Name <> "Custos" Then
		ws.Visible = xlSheetVeryHidden
	End If
Next
