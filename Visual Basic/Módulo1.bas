Attribute VB_Name = "Módulo1"
Sub ExtraerMultipleInformacion()
Attribute ExtraerMultipleInformacion.VB_ProcData.VB_Invoke_Func = "f\n14"
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim celda As Range
    Dim Texto As String
    Dim m As Object
    Dim matches As Object
    Dim i As Integer
    Dim filaDestino As Long
    ' Cambia los nombres de las hojas según tu necesidad
    Set wsOrigen = ThisWorkbook.Sheets("Mail")
    Set wsDestino = ThisWorkbook.Sheets("Datos")
    ' Inicializa la fila de destino en la hoja de destino
    filaDestino = 2
    ' Recorre cada celda en la columna B de la hoja de origen
    For Each celda In wsOrigen.Range("B1:B" & wsOrigen.Cells(wsOrigen.Rows.Count, "B").End(xlUp).Row)
        Texto = celda.Value
        ' Usar expresión regular para encontrar números de 8 dígitos
        Set m = CreateObject("VBScript.RegExp")
        With m
            .Pattern = "\b\d{8}\b"
            .Global = True
            If .Test(Texto) Then
                Set matches = .Execute(Texto)
                For i = 0 To matches.Count - 1
                    ' Pegar el valor como número
                    wsDestino.Cells(filaDestino, 1).Value2 = CLng(matches(i).Value)
                    filaDestino = filaDestino + 1
                Next i
            End If
        End With
    Next celda
MsgBox "Macro Ejecutada"
End Sub
