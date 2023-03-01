Attribute VB_Name = "Módulo1"
Sub Compilado1()
Attribute Compilado1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Compilado1 Macro
'

Dim Contador As Integer
Dim Cantidad As Integer
Dim Archivo As String
Dim CerrarArchivo As String

Contador = 1
Sheets("Listado").Select
Cantidad = Range("B1")

Do While Contador <= Cantidad

Sheets("Listado").Select
Archivo = Cells(Contador, 1)
CerrarArchivo = Cells(Contador, 3)

Workbooks.OpenText Filename:=Archivo

Windows(CerrarArchivo).Activate
    Range("A1:AZ171").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1:AZ1048575").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1:AZ1250").Select
    Selection.Copy
    Windows("Compilacion.xlsm").Activate
    
    Sheets("Compilacion").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1048575").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
    Range("B1091").Select
    



ThisWorkbook.Activate
Sheets("Listado").Select


Workbooks(CerrarArchivo).Close SaveChanges:=False



Contador = Contador + 1

Loop


End Sub
