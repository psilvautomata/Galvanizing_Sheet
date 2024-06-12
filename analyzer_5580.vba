Sub analyzer_5580()

Application.ScreenUpdating = False 'Disable screen updating
Application.EnableEvents = False 'Disable excel events

Dim i, o, Min, Max As Integer 'Variable declarations

Workbooks.Open Filename:="Path\BD_Certificados.xlsm" 'File path

Workbooks("Galvanização.xlsm").Activate 'Activate Excel Workbook
Sheets("N 5580").Activate 'Activate Excel Sheet
Range("C6:I75").ClearContents 'Clear range contents

Range("B1").Select
Range(Selection, Selection.End(xlDown)).Select 'Goes to the last empty row
Min = Selection.Rows.Count + 1 'Count the first row of Table
Max = Range("Tabela4").Rows.Count + Min - 1 'Count the last line of Table 4
Range("C6:I75").ClearContents 'Clear range contents
Range("B" & Min).Select 'Select the first Table row

For i = Min To Max 'Loop in table range

    If Range("B" & i).Value <> "" Then
    
        Workbooks("Galvanização.xlsm").Activate
        varLote = Range("B" & i).Value
        
        Workbooks("BD_Certificados.xlsm").Activate 'Activates the Workbook from DB
        Worksheets("Dados_galv").Activate
        Worksheets("Dados_galv").Range("A2").Value = varLote
        
        'Copies core values from DB Workbook
        C = Worksheets("Dados_galv").Range("B2").Value
        Si = Worksheets("Dados_galv").Range("C2").Value
        Mn = Worksheets("Dados_galv").Range("D2").Value
        P = Worksheets("Dados_galv").Range("E2").Value
        S = Worksheets("Dados_galv").Range("F2").Value
        Acabamento = Worksheets("Dados_galv").Range("T2").Value
        Mat = Worksheets("Dados_galv").Range("S2").Value
        
        'Paste the core values in Workbook "Galvanizção"
        Workbooks("Galvanização.xlsm").Activate
        Worksheets("N 5580").Range("E" & i).Value = C
        Worksheets("N 5580").Range("F" & i).Value = Mn
        Worksheets("N 5580").Range("G" & i).Value = Si
        Worksheets("N 5580").Range("H" & i).Value = P
        Worksheets("N 5580").Range("I" & i).Value = S
        Worksheets("N 5580").Range("C" & i).Value = Mat
        Worksheets("N 5580").Range("D" & i).Value = Acabamento

    End If

Next

Workbooks("BD_Certificados.xlsm").Close SaveChanges:=False 'Close BD workbook withou saving

Application.ScreenUpdating = True 'Enable screen updating
Application.EnableEvents = True 'Enable excel events

MsgBox ("Dados importados com sucesso!")


End Sub
