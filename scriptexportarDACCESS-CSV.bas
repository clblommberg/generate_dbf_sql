Sub ExportTableDesign()

    Dim db As Database, td As TableDef
    Dim filePath As String, txt As String
    
    Set db = CurrentDb()
    
    filePath = "C:\Users\claudio\Documents\ExportaConcar\Exportadesing\"
  
    For Each td In db.TableDefs
  
      If Left(td.Name, 4) <> "MSys" Then
  
        txt = "TABLA: " & td.Name & vbCrLf & vbCrLf
  
        For Each fld In td.Fields
        
          If fld.Type = dbText Then
            txt = txt & fld.Name & " TEXT(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbInteger Then
            txt = txt & fld.Name & " INTEGER(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbLong Then
            txt = txt & fld.Name & " LONG(" & fld.Type & ") " & vbCrLf
          ElseIf fld.Type = dbSingle Then
            txt = txt & fld.Name & " SINGLE(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbDouble Then
            txt = txt & fld.Name & " DOUBLE(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbCurrency Then
            txt = txt & fld.Name & " CURRENCY(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbAutoNumber Then
            txt = txt & fld.Name & " AUTONUMBER(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbDate Then
            txt = txt & fld.Name & " DATE" & vbCrLf
          ElseIf fld.Type = dbBoolean Then
            txt = txt & fld.Name & " BOOLEAN(" & fld.Type & ")" & vbCrLf
          ElseIf fld.Type = dbMemo Then
            txt = txt & fld.Name & " MEMO(" & fld.Type & ")" & vbCrLf
          End If
          
        Next fld
  
        txt = txt & vbCrLf
  
        'Guardar texto en archivo CSV
        Dim fnum As Integer
        fnum = FreeFile()
  
        Open filePath & td.Name & ".csv" For Output As fnum
        Print #fnum, txt
        Close fnum
  
      End If
  
    Next td
  
    Set db = Nothing
  
  End Sub
  
  