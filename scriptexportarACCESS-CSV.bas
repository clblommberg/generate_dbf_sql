'ANTES DE EJECUTAR LAS MACROS SE DEBEN AGREGAR DIRECTORIO TXT A LOS ORIGENES DE DATOS ODBC

Sub ExportToCSV()

    Dim db As Database, td As TableDef
    Dim errDict As Object, tableName As String
    Dim filePath As String
    
    Set db = CurrentDb()
    Set errDict = CreateObject("Scripting.Dictionary")
    
    filePath = "C:\Users\claudio\Documents\ExportaConcar\"
    
    For Each td In db.TableDefs
      If Left(td.Name, 4) <> "MSys" Then
      
        tableName = td.Name
        
        On Error Resume Next
          DoCmd.TransferText acExportDelim, , td.Name, filePath & tableName & ".csv"
          
        If Err.Number <> 0 Then
          errDict.Add tableName, Err.Description
          Err.Clear
        End If
        
        On Error GoTo 0
      
      End If
    Next td
    
    If errDict.Count > 0 Then
      MsgBox "No se pudieron exportar las siguientes tablas: " & vbCrLf & vbCrLf & Join(errDict.Keys, ", ")
    Else
      MsgBox "¡Exportación completada!"
    End If
    
    Set db = Nothing
    Set errDict = Nothing
    
  End Sub
  