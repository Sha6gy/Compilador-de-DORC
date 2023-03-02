Sub Compilado1()
'
' Compilado1 Macro
'

Dim Contador As Integer
Dim Cantidad As Integer
Dim Archivo As String
Dim CerrarArchivo As String
Dim LastRow As Long
Dim ContRow As Long

Contador = 1
ContRow = 1
Sheets("Listado").Select
Cantidad = Range("B1")
Archivo = Cells(Contador, 1)
CerrarArchivo = Cells(Contador, 3)

Do While Contador <= Cantidad
       
    Workbooks.OpenText Filename:=Archivo
    Windows(CerrarArchivo).Activate
    Range("A1:AZ1090").Select
    Selection.Copy

    ThisWorkbook.Activate
    Sheets("Compilacion").Select
    Windows("Compilacion.xlsm").Activate
    Cells(ContRow, 1).Select
    ActiveSheet.Paste
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Application.CutCopyMode = False
    Workbooks(CerrarArchivo).Close SaveChanges:=False

    ContRow = 1 + LastRow
    Contador = Contador + 1
    
    Sheets("Listado").Select
    Archivo = Cells(Contador, 1)
    CerrarArchivo = Cells(Contador, 3)
    

Loop


End Sub
