Sub FuncImportar()
    Dim db As Database
    Dim rst As Recordset
    Dim mySql As String
    Dim dbfPath As String
    Dim tableName As Stringpip

    ' Definimos las variables
    mySql = "SELECT * FROM modelo" ' Asegúrate que la tabla se llame "modelo"
    dbfPath = "C:\Users\claudio\Desktop\Concar80\" ' Ruta donde se encuentran los archivos DBF

    Set db = CurrentDb
    Set rst = db.OpenRecordset(mySql, dbOpenSnapshot)

    ' Operamos con el Recordset
    rst.MoveFirst

    Do Until rst.EOF
        tableName = rst!modelo ' Nombre de la tabla a crear en Access
        DoCmd.TransferDatabase acImport, "dBase IV", dbfPath, acTable, rst!nombreTabla, tableName
        rst.MoveNext
    Loop

    ' Limpiamos memoria
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If

    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Sub

