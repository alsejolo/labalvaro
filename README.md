Option Compare Database

Private Sub btnActualizar_Click()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConn As String
    Dim ODBCTableName As String
    Dim AccessTableName As String
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' Establecer la conexión a la base de datos actual
    'Set db = CurrentDb()
    
    ' Especificar los nombres de las tablas
    ODBCTableName = "Tarea"
    AccessTableName = "tblUsuarios"
    
    ' Configurar la conexión ODBC
    strConn = "DRIVER={SQLite3 ODBC Driver};Database=C:\Tareas\DataBase\Tarea.db;Timeout=1000;SyncPragma=NORMAL;"
    
    Set db = CurrentDb()
    
    db.TableDefs.Refresh
    
    Set rs = CurrentDb.OpenRecordset("tblUsuarios", dbOpenDynaset)
    
     ' Añade un nuevo registro
    rs.Edit
    Form_frmUsuarios.txtNombre.SetFocus
    rs!nombre = Form_frmUsuarios.txtNombre.Text
    Form_frmUsuarios.txtApellido.SetFocus
    rs!apellido = Form_frmUsuarios.txtApellido.Text
    
    ' Guarda el nuevo registro
    rs.Update
    
    ' Cerrar el recordset
    rs.Close
    
    ' Liberar los objetos
    Set rs = Nothing
    Set db = Nothing

End Sub

Private Sub btnAgregar_Click()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConn As String
    Dim ODBCTableName As String
    Dim AccessTableName As String
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' Establecer la conexión a la base de datos actual
    'Set db = CurrentDb()
    
    ' Especificar los nombres de las tablas
    ODBCTableName = "Tarea"
    AccessTableName = "tblUsuarios"
    
    ' Configurar la conexión ODBC
    strConn = "DRIVER={SQLite3 ODBC Driver};Database=C:\Tareas\DataBase\Tarea.db;Timeout=1000;SyncPragma=NORMAL;"
    
    Set db = CurrentDb()
    
    db.TableDefs.Refresh
    
    Set rs = CurrentDb.OpenRecordset("tblUsuarios", dbOpenDynaset)
    
     ' Añade un nuevo registro
    rs.AddNew
    Form_frmUsuarios.txtNombre.SetFocus
    rs!nombre = Form_frmUsuarios.txtNombre.Text
    Form_frmUsuarios.txtApellido.SetFocus
    rs!apellido = Form_frmUsuarios.txtApellido.Text
    
    ' Guarda el nuevo registro
    rs.Update
    
    ' Cerrar el recordset
    rs.Close
    
    ' Liberar los objetos
    Set rs = Nothing
    Set db = Nothing

End Sub

Private Sub btnConsulta_Click()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConn As String
    Dim ODBCTableName As String
    Dim AccessTableName As String
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' Establecer la conexión a la base de datos actual
    Set db = CurrentDb()
    
    ' Especificar los nombres de las tablas
    ODBCTableName = "Tarea"
    AccessTableName = "tblUsuarios"
    
    ' Configurar la conexión ODBC
    strConn = "DRIVER={SQLite3 ODBC Driver};Database=C:\Tareas\DataBase\Tarea.db;Timeout=1000;SyncPragma=NORMAL;"
    
    ' Crear la tabla vinculada si no existe
    On Error Resume Next
    Set tdf = db.CreateTableDef(AccessTableName, dbAttachSavePWD, ODBCTableName, strConn)
    If Err.Number = 0 Then
        db.TableDefs.Append (tdf)
    End If
    On Error GoTo 0
    
    ' Refrescar la lista de tablas
    db.TableDefs.Refresh
    
    ' Definir la consulta SQL
    strSQL = "SELECT * FROM " & AccessTableName
    
    ' Ejecutar la consulta
    Set rs = db.OpenRecordset(strSQL)
    
    ' Procesar los registros obtenidos
    Do While Not rs.EOF
        ' Realiza lo que necesites con cada registro (ejemplo, imprimir el valor del primer campo)
        Debug.Print rs.Fields(0).Value
        Form_frmUsuarios.txtNombre1.SetFocus
        Form_frmUsuarios.txtNombre1.Text = rs.Fields(1).Value
        Form_frmUsuarios.txtApellido1.SetFocus
        Form_frmUsuarios.txtApellido1.Text = rs.Fields(2).Value
        Form_frmUsuarios.txtNombre2.SetFocus
        rs.MoveNext
        Debug.Print rs.Fields(0).Value
        Form_frmUsuarios.txtNombre2.Text = rs.Fields(1).Value
        Form_frmUsuarios.txtApellido2.SetFocus
        Form_frmUsuarios.txtApellido2.Text = rs.Fields(2).Value
                
        rs.MoveNext
    Loop
    
    ' Cerrar el recordset
    rs.Close
    
    ' Liberar los objetos
    Set rs = Nothing
    Set db = Nothing
End Sub



Private Sub btnEliminar_Click()

End Sub
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConn As String
    Dim ODBCTableName As String
    Dim AccessTableName As String
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' Establecer la conexión a la base de datos actual
    'Set db = CurrentDb()
    
    ' Especificar los nombres de las tablas
    ODBCTableName = "Tarea"
    AccessTableName = "tblUsuarios"
    
    ' Configurar la conexión ODBC
    strConn = "DRIVER={SQLite3 ODBC Driver};Database=C:\Tareas\DataBase\Tarea.db;Timeout=1000;SyncPragma=NORMAL;"
    
    Set db = CurrentDb()
    
    db.TableDefs.Refresh
    
    Set rs = CurrentDb.OpenRecordset("tblUsuarios", dbOpenDynaset)
    
     ' Añade un nuevo registro
    rs.FindFirst "nombre = 'Form_frmUsuarios.txtNombre.Text'"

    If Not rs.NoMatch Then
        rs.Delete
    End If
    
    ' Guarda el nuevo registro
    rs.Update
    
    ' Cerrar el recordset
    rs.Close
    
    ' Liberar los objetos
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Sub btnLimpiar_Click()
Form_frmUsuarios.txtNombre.SetFocus
Form_frmUsuarios.txtNombre.Text = ""
Form_frmUsuarios.txtApellido.SetFocus
Form_frmUsuarios.txtApellido.Text = ""
Form_frmUsuarios.txtNombre1.SetFocus
Form_frmUsuarios.txtNombre1.Text = ""
Form_frmUsuarios.txtApellido1.SetFocus
Form_frmUsuarios.txtApellido1.Text = ""
Form_frmUsuarios.txtNombre2.SetFocus
Form_frmUsuarios.txtNombre2.Text = ""
Form_frmUsuarios.txtApellido2.SetFocus
Form_frmUsuarios.txtApellido2.Text = ""
Form_frmUsuarios.txtNombre.SetFocus
End Sub


Private Sub Comando71_Click()

End Sub

Private Sub Form_Load()
Form_frmUsuarios.txtNombre.SetFocus
Form_frmUsuarios.txtNombre.Text = ""
Form_frmUsuarios.txtApellido.SetFocus
Form_frmUsuarios.txtApellido.Text = ""
Form_frmUsuarios.txtNombre1.SetFocus
Form_frmUsuarios.txtNombre1.Text = ""
Form_frmUsuarios.txtApellido1.SetFocus
Form_frmUsuarios.txtApellido1.Text = ""
Form_frmUsuarios.txtNombre2.SetFocus
Form_frmUsuarios.txtNombre2.Text = ""
Form_frmUsuarios.txtApellido2.SetFocus
Form_frmUsuarios.txtApellido2.Text = ""
Form_frmUsuarios.txtNombre.SetFocus
End Sub

Private Sub txtApellido1_DblClick(Cancel As Integer)
Form_frmUsuarios.txtApellido.SetFocus
Form_frmUsuarios.txtApellido.Text = Form_frmUsuarios.txtApellido1.Value
End Sub

Private Sub txtApellido2_DblClick(Cancel As Integer)
Form_frmUsuarios.txtApellido.SetFocus
Form_frmUsuarios.txtApellido.Text = Form_frmUsuarios.txtApellido2.Value
End Sub

Private Sub txtNombre1_DblClick(Cancel As Integer)
Form_frmUsuarios.txtNombre.SetFocus
Form_frmUsuarios.txtNombre.Text = Form_frmUsuarios.txtNombre1.Value
End Sub

Private Sub txtNombre2_DblClick(Cancel As Integer)
Form_frmUsuarios.txtNombre.SetFocus
Form_frmUsuarios.txtNombre.Text = Form_frmUsuarios.txtNombre2.Value
End Sub
