Option Explicit

Public Sub WriteLog(ByVal mensaje As String, Optional ByVal nivel As String = "INFO")
    Dim FSO As Object
    Dim archivo As Object
    Dim rutaArchivo As String
    Dim fechaHora As String
    
    ' Definir la ruta del archivo de log
    rutaArchivo = ThisWorkbook.Path & "\log.txt"
    
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Abrir archivo en modo Append (si no existe, se crea)
    Set archivo = FSO.OpenTextFile(rutaArchivo, 8, True)
    
    ' Obtener fecha y hora actual
    fechaHora = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Escribir mensaje con formato
    archivo.WriteLine fechaHora & " [" & nivel & "] " & mensaje
    
    ' Cerrar archivo
    archivo.Close
    
    ' Liberar objetos
    Set archivo = Nothing
    Set FSO = Nothing
End Sub
