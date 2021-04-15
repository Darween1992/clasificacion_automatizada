Attribute VB_Name = "Biblioteca"
Sub MOSTRAR_OCULTAS()

For Each Sheet In ActiveWorkbook.Sheets


Sheet.Visible = xlSheetVisible

Next Sheet


End Sub


Sub FORMULA_SEMANA()

Dim lote As String

lote = ActiveSheet.Name

Range("A4").Value = lote
Range("A5").Value = lote
Range("A6").Value = lote
Range("A7").Value = lote
Range("A8").Value = lote
Range("A9").Value = lote
Range("A10").Value = lote
Range("A11").Value = lote
Range("A12").Value = lote
Range("A13").Value = lote
Range("A14").Value = lote
Range("A15").Value = lote
Range("A16").Value = lote
Range("A17").Value = lote
Range("A18").Value = lote
Range("A19").Value = lote
Range("A20").Value = lote
Range("A21").Value = lote
Range("A22").Value = lote
Range("A23").Value = lote
Range("A24").Value = lote
Range("A25").Value = lote
Range("A26").Value = lote
Range("A27").Value = lote
Range("A28").Value = lote
Range("A29").Value = lote
Range("A30").Value = lote
Range("A31").Value = lote
Range("A32").Value = lote
Range("A33").Value = lote
Range("A34").Value = lote

'FORMULA CONDICIONAL PARA CONCATENAR Y CREAR CODIGO DE VERIFICACION

    Range("AI4").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-31]<>"""",CONCATENATE(RC[-32],RC[-34]),"""")"
    Range("AI4").Select
    Selection.AutoFill Destination:=Range("AI4:AI34"), Type:=xlFillDefault
    Range("AI4:AI34").Select
    Range("AI28").Select

'FORMULA BUSCAR V PARA SACAR SEMANA DE PRODUCCION


 Range("B4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[33],consolidado_dinamico.xlsm!Tabla3[#Data],4,FALSE)"
    Range("B4").Select
    Selection.AutoFill Destination:=Range("B4:B34"), Type:=xlFillDefault
    Range("B4:B34").Select
    ActiveWindow.SmallScroll Down:=0
    Range("B35").Select
    
    

End Sub
Sub consolodar_informacion()

Range("A4:AI34").Select
Selection.Copy

Sheets(1).Select

Range("A1").End(xlDown).Offset(1, 0).Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

End Sub



Sub Quita_Duplicados()


Range("AG2").Select
Do While Not IsEmpty(ActiveCell)
    x = WorksheetFunction.CountIf(Range("AG:AG"), ActiveCell)
    If x > 1 Then
        ActiveCell.EntireRow.Delete
    Else
        ActiveCell.Offset(1, 0).Select
    End If
Loop
Range("A1").Select
End Sub


Sub BUSCARV_DOBLE()
'
' BUSCARV_DOBLE Macro
'

'
    
    Range("AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=21
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AH2:AH7193").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AH2:AH1048574").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-18
    Range("AH6872").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@codigo],Tabla3[#All],5,FALSE)"
    Range("AH6872").Select
    Selection.Copy
    Range("AI6872").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@codigo],Tabla3[#All],6,FALSE)"
    Range("AH6872:AI6872").Select
    Selection.AutoFill Destination:=Range("AH6872:AI7196")
    Range("AH6872:AI7196").Select
        Range("AH7196").Select
    ActiveWindow.SmallScroll Down:=6
End Sub


Sub QUITAR_FILTROS()

Sheets("consolidado1").Select

    ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=5
        ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=34
    ActiveSheet.ListObjects("Tabla2").Range.AutoFilter Field:=35
        Sheets("Dashboart").Select
    Range("G3").Select
End Sub

