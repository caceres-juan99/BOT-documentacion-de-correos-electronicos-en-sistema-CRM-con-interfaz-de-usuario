Attribute VB_Name = "Módulo2"
Sub BuscarYCopiarOTPs()
Attribute BuscarYCopiarOTPs.VB_ProcData.VB_Invoke_Func = "n\n14"
    Dim wsDatos As Worksheet, wsMail As Worksheet
    Dim rngBusqueda As Range, celdaBusqueda As Range
    Dim rngBusquedaVAL_OTP As Range, celdaVAL_OTP As Range
    Dim j As Long
    Dim encontrado As Boolean
    ' Definir las hojas de trabajo
    Set wsDatos = ThisWorkbook.Sheets("Datos") ' Hoja Datos
    Set wsMail = ThisWorkbook.Sheets("Mail") ' Hoja Mail
    ' Definir los rangos de búsqueda
    Set rngBusqueda = wsMail.Range("B2:B" & wsMail.Cells(wsMail.Rows.Count, "B").End(xlUp).Row) ' Rango de la columna B en Hoja Mail
    Set rngBusquedaVAL_OTP = wsDatos.Range("B2:B" & wsDatos.Cells(wsDatos.Rows.Count, "B").End(xlUp).Row) ' Rango de la columna B en Hoja Datos
    ' Iterar sobre cada celda en VAL_OTP (columna B de Datos)
    For Each celdaVAL_OTP In rngBusquedaVAL_OTP
        encontrado = False
        ' Iterar sobre cada celda en Subject (columna B de Mail)
        For Each celdaBusqueda In rngBusqueda
            If InStr(1, celdaBusqueda.Value, celdaVAL_OTP.Value) > 0 Then ' Verifica si VAL_OTP está contenido en Subject
                encontrado = True
                ' Copiar la fila encontrada en Hoja Mail a partir de la columna D en Hoja Datos
                For j = 1 To celdaBusqueda.EntireRow.Cells.Count
                    ' Asegurar que la columna no exceda el límite de columnas de Excel
                    If (4 + j - 1) <= wsDatos.Columns.Count Then
                        wsDatos.Cells(celdaVAL_OTP.Row, 4 + j - 1).Value = celdaBusqueda.EntireRow.Cells(1, j).Value
                    Else
                        Exit For ' Salir del bucle si excedemos el límite de columnas
                    End If
                Next j
                Exit For
            End If
        Next celdaBusqueda
        ' Si no se encontró ninguna coincidencia, opcionalmente se puede dejar en blanco o poner un mensaje
        If Not encontrado Then
            wsDatos.Cells(celdaVAL_OTP.Row, 4).Value = "No encontrado"
        End If
    Next celdaVAL_OTP
MsgBox "Macro Ejecutada"
End Sub
