Attribute VB_Name = "M�dulo3"
Sub ActualizarConexiones()
Attribute ActualizarConexiones.VB_ProcData.VB_Invoke_Func = "h\n14"
    Dim conexion As WorkbookConnection
    Dim totalConexiones As Integer
    Dim conexionesExitosas As Integer
    ' Inicializar contadores
    totalConexiones = 0
    conexionesExitosas = 0
    ' Contar las conexiones existentes
    totalConexiones = ThisWorkbook.Connections.Count
    ' Si no hay conexiones, salir y notificar al usuario
    If totalConexiones = 0 Then
        MsgBox "No hay conexiones para actualizar.", vbExclamation, "Sin Conexiones"
        Exit Sub
    End If
    ' Iterar por cada conexi�n y actualizarla
    For Each conexion In ThisWorkbook.Connections
        On Error Resume Next ' Manejar posibles errores de actualizaci�n
        conexion.Refresh ' Actualizar la conexi�n
        DoEvents ' Permitir que Excel procese otras tareas mientras espera
        If Err.Number = 0 Then
            conexionesExitosas = conexionesExitosas + 1
        End If
        On Error GoTo 0 ' Restaurar manejo de errores predeterminado
    Next conexion
    
    ' Esperar hasta que todas las conexiones terminen (modo s�ncrono)
    Application.CalculateUntilAsyncQueriesDone
    
    ' Mostrar mensaje de finalizaci�n en el formulario
    If conexionesExitosas = totalConexiones Then
        ' Muestra el formulario con el mensaje de �xito
        Proceso_completado.Texto.Caption = "�Actualizaci�n exitosa!"
        Proceso_completado.Show
    Else
        MsgBox "Actualizaci�n completada con errores. Verifica las conexiones.", vbExclamation, "Proceso Incompleto"
    End If
End Sub
