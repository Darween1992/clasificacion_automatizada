Attribute VB_Name = "Index"
Sub principal()


Dim respuesta As Byte
Dim titulo As String

titulo = "Mensaje de Aviso"
respuesta = MsgBox("Se actualizara la base de datos 'consolidado1' ?" & Chr(13) & vbNewLine & " ¿Desea Continuar?", vbQuestion + vbYesNo, titulo)

If respuesta = vbYes Then


Dim CLASIFICACION2020 As Workbook
Set CLASIFICACION2020 = Workbooks.Open("C:\DATOS\TRABAJO\REPORTE DIARIO\Datos Diarios\CLASIFICACION HUEVO OPAV\CLASIFICACION OPAV.xlsx")

Call MOSTRAR_OCULTAS

Dim I As Byte


Dim N_HOJAS  As Integer

N_HOJAS = Sheets.Count

'MsgBox N_HOJAS


For I = 1 To N_HOJAS

If Sheets(I).Range("A3").Value = "LOTE" Then

Sheets(I).Select

Call FORMULA_SEMANA

End If

Next I

ActiveWorkbook.Sheets.Add before:=Sheets(1)

Range("A1").Value = "lote"
Range("A2").Value = "SIN DATOS"

'segundo bucle

For I = 2 To N_HOJAS

If Sheets(I).Range("A3").Value = "LOTE" Then

Sheets(I).Select



Call consolodar_informacion

End If

Next I

'FITRADO DE INFORMACION CONSOLIDADA

Sheets(1).Select

 Columns("AG:AH").Select
    Selection.Delete Shift:=xlToLeft
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "FILTRO"
    Range("AG1").Select
    Selection.AutoFilter
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
  
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.Range("$AG$1:$AG$281").AutoFilter Field:=1, Criteria1:="<>"
   
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Windows("consolidado_dinamico.xlsm").Activate
    
    Sheets("consolidado1").Select
    
    Range("A3").End(xlDown).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

CLASIFICACION2020.Close savechanges:=False

MsgBox "Quitando valores duplicados, Esto puede tardar algunos minutos", vbInformation, "Informacion"

Call Quita_Duplicados

Call BUSCARV_DOBLE

MsgBox "Informacion actualizada", vbInformation, "Fin de proceso"

Sheets("Dashboart").Select
ActiveWorkbook.RefreshAll


Else
    MsgBox "No se actualizo ningun registro de la base de datos 'consolidado1'"
End If

End Sub
