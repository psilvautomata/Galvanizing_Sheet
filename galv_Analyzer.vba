Sub galv_Analyzer()

Application.ScreenUpdating = False 'Disable screen updating
Application.EnableEvents = False 'Disable excel events

Dim i, o, Min, Max As Integer 'Variable declarations

Workbooks.Open Filename:="Path\BD_Certificados.xlsm" 'File path

Workbooks("Galvanização.xlsm").Activate 'Activate Excel Workbook
Sheets("Soufer").Activate 'Activate Excel Sheet

Range("B1").Select
Range(Selection, Selection.End(xlDown)).Select 'Goes to last empty row
Min = Selection.Rows.Count + 1 'Count the first row of Table
Max = Range("Tabela1").Rows.Count + Min - 1 'Count the last row of Table
Range("C6:F75").ClearContents 'Clear range contents
Range("B" & Min).Select 'Select the first Table row

For i = Min To Max 'Loop in table range

    If Range("B" & i).Value <> "" Then
    
        Workbooks("Galvanização.xlsm").Activate 'Activate the Workbook from DB
        varLote = Range("B" & i).Value
        
        Workbooks("BD_Certificados.xlsm").Activate
        Worksheets("Dados_galv").Activate
        Worksheets("Dados_galv").Range("A2").Value = varLote
        
        'Copies core values from DB Workbook
        Si = Worksheets("Dados_galv").Range("C2").Value
        P = Worksheets("Dados_galv").Range("E2").Value
        Acabamento = Worksheets("Dados_galv").Range("T2").Value
        Mat = Worksheets("Dados_galv").Range("S2").Value
        
        'Pastes core values from DB Workbook
        Workbooks("Galvanização.xlsm").Activate
        Worksheets("Soufer").Range("C" & i).Value = Mat
        Worksheets("Soufer").Range("D" & i).Value = Acabamento
        Worksheets("Soufer").Range("E" & i).Value = Si
        Worksheets("Soufer").Range("F" & i).Value = P

    End If

Next

Workbooks("BD_Certificados.xlsm").Close SaveChanges:=False 'Close BD workbook withou saving


Application.ScreenUpdating = True 'Enable screen updating
Application.EnableEvents = True 'Enable excel events

'MsgBox ("Dados importados com sucesso!")


End Sub
