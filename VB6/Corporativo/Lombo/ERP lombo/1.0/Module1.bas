Attribute VB_Name = "Module1"
'*********************************************************************************************
'Nombre: Punto de venta
'Proposito: Manejo de Compras, Ventas e inventarios
'
'Version    Fecha       Nombre  Descripcion
'---------------------------------------------------------------------------------------------
'1.0        13/02/2020  JAHG    Creacion del programa
'1.1        14/02/2020  JAHG    Definicion de tablas de usuario
'1.2        15/02/2020  JAHG    Leer datos de conexion a la base de datos desde archivo
'1.3        17/02/2020  JAHG    Validacion de existencia de usuario, base de datos y tablas
'1.4        17/02/2020  JAHG    Creacion del cuadro de inicio de sesion
'1.5        18/02/2020  JAHG    Creacion del menu principal
'1.6        18/02/2020  JAHG    Creacion del submenu
'1.7        19/02/2020  JAHG    Creacion de la pantalla Grupo de concurrentes
'1.8        20/02/2020  JAHG    Creacion del buscador
'1.9        21/02/2020  JAHG    Creacion del menu alta de cajas
'1.10       24/02/2020  JAHG    Creacion del menu alta de usuarios
'1.11       29/02/2020  JAHG    Creacion del menu alta de reportes
'1.12       02/03/2020  JAHG    Creacion del menu ejecutar reporte
'1.13       03/03/2020  JAHG    Creacion de la lista categorias
'1.14       04/03/2020  JAHG    Creacion de la lista unidades de medida
'1.15       05/03/2020  JAHG    Creacion de la lista unidades de tipos de transaccion
'1.16       05/03/2020  JAHG    Creacion de la lista unidades de subinventarios
'1.17       05/03/2020  JAHG    Creacion del menu atributos
'1.18       05/03/2020  JAHG    Creacion del menu articulos
'**********************************************************************************************




'==============================================================================================

'V  A   R   I   A   B   L   E   S

'==============================================================================================
'Variables de conexion inicial a la base de datos
Global Instancia As String
Global CnString As String
Global Cn As New ADODB.Connection
'Variables de inicio de sesion
Global RsUser As New ADODB.Recordset
Global StUser As String
Global InUserId As Integer
Global ErrorCount As Integer
'Variables de Menu Inicial
Global RsMenuInicial As New ADODB.Recordset
Global StMenuInicial As String
Global InRespId As Integer
'Variables de subMenu
Global RsSubMenu As New ADODB.Recordset
Global SubMenuRs1 As New ADODB.Recordset
Global SubMenuRs2 As New ADODB.Recordset
Global StSubMenu As String
Global SubMenuFrameName As String
Global SubMenuSt1 As String
Global SubMenuSt2 As String
'variables del menu usarios
Global InUusario As Integer
'variables menu articulos
Global StTipoBuscador As String
Global StTipoBuscadorArticulo As String
Global StTipoGuardarArticulo As String
Option Explicit




'==============================================================================================

'F  U   N   C   I   O   N   E   S

'==============================================================================================
'Funcion para obtener el nombre de la instancia
Function Getinstancia() As String
    Dim n_File As Integer
    Dim Linea As String
    n_File = FreeFile
    Open App.Path & "\instancia" For Input As n_File
    Do While Not EOF(n_File)
        Line Input #n_File, Linea
    Loop
    Close n_File
    Getinstancia = Linea
End Function
'Funcion para obtener el password de la base de datos
Function Getpassword() As String
    Dim n_File As Integer
    Dim Linea As String
    n_File = FreeFile
    Open App.Path & "\password" For Input As n_File
    Do While Not EOF(n_File)
        Line Input #n_File, Linea
    Loop
    Close n_File
    Getpassword = Linea
End Function
'Funcion para obtener el usuario de la base de datos
Function Getusuario() As String
    Dim n_File As Integer
    Dim Linea As String
    n_File = FreeFile
    Open App.Path & "\usuario" For Input As n_File
    Do While Not EOF(n_File)
        Line Input #n_File, Linea
    Loop
    Close n_File
    Getusuario = Linea
End Function
'Funcion para obtener el id de la responsabilidad
Function GetRespId(Pdesciption As String) As Integer
    Dim RsRecorSet As New ADODB.Recordset
    Dim stString As String
    stString = "SELECT t1.responsibility_id                     " & _
               "  FROM fnd_responsibility t1                    " & _
               " WHERE t1.end_date is null                      " & _
               "   AND t1.description = '" & Pdesciption & "';  "
    With RsRecorSet
        If .State = 1 Then .Close
            .Open stString, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    GetRespId = RsRecorSet.Fields("responsibility_id")
    RsRecorSet.Close
End Function
'Funcion para obtener el nombre del frame
Function GetSubMenuFrame(Pdesciption As String, Presponsibility_id As Integer) As String
    Dim RsRecorSet As New ADODB.Recordset
    Dim stString As String
    stString = "SELECT t1.frame_name                                    " & _
               "  FROM fnd_responsibility_menu t1                       " & _
               " WHERE t1.end_date is null                              " & _
               "   AND t1.responsibility_id = " & Presponsibility_id & _
               "   AND t1.description = '" & Pdesciption & "';          "
    With RsRecorSet
        If .State = 1 Then .Close
            .Open stString, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    GetSubMenuFrame = RsRecorSet.Fields("frame_name")
    RsRecorSet.Close
End Function
'Funcion para obtener el request_unit_name
Function Getrequest_unit_name(Pdesciption As String) As String
    Dim RsRecorSet As New ADODB.Recordset
    Dim stString As String
    stString = "SELECT t1.request_unit_name                     " & _
               "  FROM fnd_request_headers t1                   " & _
               " WHERE t1.description = '" & Pdesciption & "';  "
    With RsRecorSet
        If .State = 1 Then .Close
            .Open stString, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    Getrequest_unit_name = RsRecorSet.Fields("request_unit_name")
    RsRecorSet.Close
End Function
'Funcion para obtener el caja_id
Function Getcaja_id(Pdesciption As String) As Integer
    On Error GoTo err
    Dim RsRecorSet As New ADODB.Recordset
    Dim stString As String
    stString = "SELECT caja_id                              " & _
               "  FROM fnd_cajas                            " & _
               " WHERE description = '" & Pdesciption & "'; "
    With RsRecorSet
        If .State = 1 Then .Close
            .Open stString, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    Getcaja_id = RsRecorSet.Fields("caja_id")
    RsRecorSet.Close
    Exit Function
err:
    Getcaja_id = 1
End Function
'Funcion para obtener el category_id
Function Getcategory_id(Pdesciption As String) As Integer
    On Error GoTo err
    Dim RsRecorSet As New ADODB.Recordset
    Dim stString As String
    stString = "SELECT category_id                          " & _
               "  FROM mtl_item_categories                  " & _
               " WHERE description = '" & Pdesciption & "'; "
    With RsRecorSet
        If .State = 1 Then .Close
            .Open stString, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    Getcategory_id = RsRecorSet.Fields("category_id")
    RsRecorSet.Close
    Exit Function
err:
    Getcategory_id = 61
End Function




'==============================================================================================

'B  A   S   E       D   E       D   A   T   O   S

'==============================================================================================
'Modulo para abrir la conexion a la base de datos
Sub OpenBd()

    '------------------------------------------------------------------------------------------
    'VARIABLES
    '------------------------------------------------------------------------------------------
    'variables de conexion inicial
    Dim CnInitial As New ADODB.Connection
    Dim PasswordInitial As String
    Dim UsuarioInitial As String
    Dim CnStringInitial As String
    'variables validacion de usuario
    Dim UserExists As New ADODB.Recordset
    Dim StUserExists As String
    Dim CreateUser As New ADODB.Command
    'variables validacion de BD
    Dim DBExists As New ADODB.Recordset
    Dim StDBExists As String
    Dim CreateDB As New ADODB.Command
    'variables validacion de tablas
    Dim TableExists As New ADODB.Recordset
    Dim StTableExists As String
    Dim CreateTable As New ADODB.Command
    Dim InsertIntoTable As New ADODB.Command
    'Tomamos valores para la conexion inicial
    Instancia = Getinstancia
    PasswordInitial = Getpassword
    UsuarioInitial = Getusuario
    CnStringInitial = "Provider=SQLOLEDB.1;                 " & _
                      "Password=" & PasswordInitial & ";    " & _
                      "Persist Security Info=True;          " & _
                      "User ID=" & UsuarioInitial & ";      " & _
                      "Initial Catalog=master;              " & _
                      "Data Source=" & Instancia
    'abrimos la conexion
    With CnInitial
        .CursorLocation = adUseClient
        .Open CnStringInitial
        
    End With
    
    '------------------------------------------------------------------------------------------
    'CREACION DE USUARIO DE BASE DE DATOS
    '------------------------------------------------------------------------------------------
    'validamos la existencia del usuario
    StUserExists = "SELECT COUNT(*) as usuario  " & _
                   "FROM master.sys.sql_logins  " & _
                   "WHERE name = 'erp';       "
    With UserExists
        If .State = 1 Then .Close
            .Open StUserExists, CnInitial, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    'si el usuario no existe lo creamos
    If UserExists.Fields("usuario") = 0 Then
        With CreateUser
            .CommandText = "CREATE LOGIN erp WITH PASSWORD=N'Jahg1991', CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'sysadmin'         " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'securityadmin'    " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'serveradmin'      " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'setupadmin'       " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'processadmin'     " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'diskadmin'        " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'dbcreator'        " & _
                             "EXEC sys.sp_addsrvrolemember @loginame = N'erp', @rolename = N'bulkadmin'        "
            .ActiveConnection = CnInitial
            .Execute
        End With
        UserExists.Close
    End If

    '------------------------------------------------------------------------------------------
    'CREACION DE BASE DE DATOS
    '------------------------------------------------------------------------------------------
    'validamos la existencia de la base de datos
    StDBExists = "SELECT COUNT(*)as bd          " & _
                 "FROM master.dbo.sysdatabases  " & _
                 "WHERE name = 'erp';           "
    With DBExists
        If .State = 1 Then .Close
            .Open StDBExists, CnInitial, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    'si la base de datos no existe la creamos
    If DBExists.Fields("bd") = 0 Then
        CreateDB.CommandText = "CREATE DATABASE erp"
        CreateDB.ActiveConnection = CnInitial
        CreateDB.Execute
        DBExists.Close
    End If
    With CnInitial
        .Close
    End With
    'Tomamos valores para la conexion
    CnString = "Provider=SQLOLEDB.1;        " & _
               "Password=Jahg1991;          " & _
               "Persist Security Info=True; " & _
               "User ID=erp;                " & _
               "Initial Catalog=erp;        " & _
               "Data Source=" & Instancia
    'abrimos la conexion a erp
    With Cn
        .CursorLocation = adUseClient
        .Open CnString
    End With

    '------------------------------------------------------------------------------------------
    'CREACION DE TABLAS
    '------------------------------------------------------------------------------------------
    'Validamos la existencia de las tablas
    
        '======================================================================================
        'ADMINISTRADOR DEL SISTEMA
        '======================================================================================
    
            '..................................................................................
            'fnd_user
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla           " & _
                            "FROM information_schema.tables     " & _
                            "WHERE table_catalog = 'erp'        " & _
                            "AND table_name = 'fnd_user';       "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_user(" & _
                                            "user_id                 INT             IDENTITY(1,1) PRIMARY KEY, " & _
                                            "user_name               VARCHAR(100)    NOT NULL,                  " & _
                                            "last_update_date        DATE            NOT NULL,                  " & _
                                            "last_updated_by         INT             NOT NULL,                  " & _
                                            "creation_date           DATE            NOT NULL,                  " & _
                                            "created_by              INT             NOT NULL,                  " & _
                                            "last_update_login       INT,                                       " & _
                                            "encrypted_user_password VARCHAR(100)    NOT NULL,                  " & _
                                            "start_date              DATE            NOT NULL,                  " & _
                                            "end_date                DATE,                                      " & _
                                            "description             VARCHAR(240),                              " & _
                                            "last_logon_date         DATE,                                      " & _
                                            "password_date           DATE,                                      " & _
                                            "caja                    VARCHAR(100)                               " & _
                                          ");                                                                   "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user             " & _
                                                "(user_name,                    " & _
                                                "last_update_date,              " & _
                                                "last_updated_by,               " & _
                                                "creation_date,                 " & _
                                                "created_by,                    " & _
                                                "last_update_login,             " & _
                                                "encrypted_user_password,       " & _
                                                "start_date,                    " & _
                                                "description)                   " & _
                                              "Values                           " & _
                                                "('sysadmin',                   " & _
                                                "GETDATE(),                     " & _
                                                "1,                             " & _
                                                "GETDATE(),                     " & _
                                                "1,                             " & _
                                                "1,                             " & _
                                                "'sysadmin',                    " & _
                                                "GETDATE(),                     " & _
                                                "'Administrador del Sistema');  "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'fnd_responsibility
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla               " & _
                            "FROM information_schema.tables         " & _
                            "WHERE table_catalog = 'erp'            " & _
                            "AND table_name = 'fnd_responsibility'; "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_responsibility(                                        " & _
                                        "responsibility_id       INT             IDENTITY(1,1) PRIMARY KEY, " & _
                                        "last_update_date        DATE,                                      " & _
                                        "last_updated_by         INT,                                       " & _
                                        "creation_date           DATE            NOT NULL,                  " & _
                                        "created_by              INT             NOT NULL,                  " & _
                                        "start_date              DATE            NOT NULL,                  " & _
                                        "end_date                DATE,                                      " & _
                                        "request_group_id        INT,                                       " & _
                                        "description             VARCHAR(100),                              " & _
                                   ");                                                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Administrador del Sistema');  "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "2,                             " & _
                                    "'Inventarios');                "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "3,                             " & _
                                    "'Compras');                    "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "4,                             " & _
                                    "'Ventas');                     "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "5,                             " & _
                                    "'Produccion');                 "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility  " & _
                                    "(last_update_date,             " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "start_date,                    " & _
                                    "request_group_id,              " & _
                                    "description)                   " & _
                                   "VALUES(                         " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "6,                             " & _
                                    "'Caja');                       "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
        
            '..................................................................................
            'fnd_request_groups
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla               " & _
                            "FROM information_schema.tables         " & _
                            "WHERE table_catalog = 'erp'            " & _
                            "AND table_name = 'fnd_request_groups'; "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_request_groups(                                " & _
                                    "request_group_id     INT           IDENTITY(1,1) PRIMARY KEY,  " & _
                                    "request_group_name   VARCHAR(30)   NOT NULL,                   " & _
                                    "last_update_date     DATE          NOT NULL,                   " & _
                                    "last_updated_by      INT           NOT NULL,                   " & _
                                    "creation_date        DATE          NOT NULL,                   " & _
                                    "created_by           INT           NOT NULL,                   " & _
                                    "description          VARCHAR(80)   NOT NULL                    " & _
                                   ");                                                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_SYS',                       " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Sysadmin'     " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_INV',                       " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Inventarios'  " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_PO',                        " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Compras'      " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_AR',                        " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Ventas'       " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_WIP',                       " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Produccion'   " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_groups( " & _
                                    "request_group_name,            " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "'C_CAJA',                      " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Concurrentes de Caja'         " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
        
            '..................................................................................
            'fnd_request_group_units
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                       " & _
                            "FROM information_schema.tables                 " & _
                            "WHERE table_catalog = 'erp'                    " & _
                            "AND table_name = 'fnd_request_group_units';    "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_request_group_units(                           " & _
                                    "request_group_id   INT         NOT NULL,                       " & _
                                    "request_unit_id    INT         IDENTITY(1,1)   PRIMARY KEY,    " & _
                                    "request_unit_name  VARCHAR(30) NOT NULL,                       " & _
                                    "last_update_date   DATE        NOT NULL,                       " & _
                                    "last_updated_by    INT         NOT NULL,                       " & _
                                    "creation_date      DATE        NOT NULL,                       " & _
                                    "created_by         INT         NOT NULL,                       " & _
                                    "description        VARCHAR(80)                                 " & _
                                   ");                                                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_group_units( " & _
                                    "request_group_id,              " & _
                                    "request_unit_name,             " & _
                                    "last_update_date,              " & _
                                    "last_updated_by,               " & _
                                    "creation_date,                 " & _
                                    "created_by,                    " & _
                                    "description                    " & _
                                   ")VALUES(                        " & _
                                    "1,                             " & _
                                    "'USUARIOS_ACTIVOS',            " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "GETDATE(),                     " & _
                                    "1,                             " & _
                                    "'Listado de usuarios Activos'  " & _
                                   ");                              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'fnd_user_resp_groups_direct
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                           " & _
                            "FROM information_schema.tables                     " & _
                            "WHERE table_catalog = 'erp'                  " & _
                            "AND table_name = 'fnd_user_resp_groups_direct';    "
                With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_user_resp_groups_direct(   " & _
                                    "user_id            INT,                    " & _
                                    "responsibility_id  INT,                    " & _
                                    "security_group_id  INT,                    " & _
                                    "start_date         DATE,                   " & _
                                    "end_date           DATE,                   " & _
                                    "created_by         INT,                    " & _
                                    "creation_date      DATE,                   " & _
                                    "last_updated_by    INT,                    " & _
                                    "last_update_date   DATE                    " & _
                                ");                                             "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "2,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "3,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                                                
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "4,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                                                
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "5,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                    "user_id,                                   " & _
                                    "responsibility_id,                         " & _
                                    "start_date,                                " & _
                                    "created_by,                                " & _
                                    "creation_date,                             " & _
                                    "last_updated_by,                           " & _
                                    "last_update_date                           " & _
                                   ")VALUES(                                    " & _
                                    "1,                                         " & _
                                    "6,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE(),                                 " & _
                                    "1,                                         " & _
                                    "GETDATE()                                  " & _
                                   ");                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'fnd_responsibility_menu
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla               " & _
                            "FROM information_schema.tables         " & _
                            "WHERE table_catalog = 'erp'      " & _
                            "AND table_name = 'fnd_responsibility_menu'; "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_responsibility_menu(                                   " & _
                                        "responsibility_menu_id  INT             IDENTITY(1,1) PRIMARY KEY, " & _
                                        "responsibility_id       INT             NOT NULL,                  " & _
                                        "last_update_date        DATE,                                      " & _
                                        "last_updated_by         INT,                                       " & _
                                        "creation_date           DATE            NOT NULL,                  " & _
                                        "created_by              INT             NOT NULL,                  " & _
                                        "start_date              DATE            NOT NULL,                  " & _
                                        "end_date                DATE,                                      " & _
                                        "description             VARCHAR(100),                              " & _
                                        "frame_name              VARCHAR(100)                               " & _
                                   ");                                                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Administrador del sistema
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Grupos de Concurrentes',              " & _
                                    "'FGruposConcurrentes'                  " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Cajas',                               " & _
                                    "'FCajas'                               " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Usuarios',                            " & _
                                    "'FUsuarios'                            " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Creacion de Reportes',                " & _
                                    "'FCreacionReportes'                    " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                                                
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Inventarios
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Articulos',                           " & _
                                    "'FArticulos'                           " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Consumo de insumos'                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Ajuste de inventario'                 " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Visualisar transacciones'             " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Alta de Subinventarios'               " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "2,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Compras
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Proveedores'                          " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Ordenes de compra'                    " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Recepciones'                          " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Devoluciones'                         " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Pagos'                                " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "3,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Ventas
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "4,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Clientes'                             " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "4,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Factura de venta'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "4,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Devoluciones'                         " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "4,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Cobros'                               " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "4,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Produccion
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Listas de materiales'                 " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Consumos'                             " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Devolucion de consumos'               " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Produccion'                           " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Devolucion de produccion'             " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "5,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                
                'Submenu Caja
                '______________________________________________________________________________
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "6,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Corte de caja'                        " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description                            " & _
                                   ")VALUES(                                " & _
                                    "6,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Ajuste de caja'                       " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_responsibility_menu(    " & _
                                    "responsibility_id,                     " & _
                                    "last_update_date,                      " & _
                                    "last_updated_by,                       " & _
                                    "creation_date,                         " & _
                                    "created_by,                            " & _
                                    "start_date,                            " & _
                                    "description,                           " & _
                                    "frame_name                             " & _
                                   ")VALUES(                                " & _
                                    "6,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "1,                                     " & _
                                    "GETDATE(),                             " & _
                                    "'Reportes',                            " & _
                                    "'FEjecutarReporte'                     " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'fnd_request_headers
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                   " & _
                            "FROM information_schema.tables             " & _
                            "WHERE table_catalog = 'erp'          " & _
                            "AND table_name = 'fnd_request_headers';    "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_request_headers(                                        " & _
                                   " request_header_id   int                 IDENTITY(1,1)   PRIMARY KEY,    " & _
                                   " request_unit_name   varchar(80)         NOT NULL,                       " & _
                                   " last_update_date    date                NOT NULL,                       " & _
                                   " last_updated_by     int                 NOT NULL,                       " & _
                                   " creation_date       date                NOT NULL,                       " & _
                                   " created_by          int                 NOT NULL,                       " & _
                                   " description         varchar(80)         NULL,                           " & _
                                   " query               varchar(8000)       NOT NULL,                       " & _
                                   " parametros          int                 NOT NULL);                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_request_headers                                                                                                                            " & _
                                   "(request_unit_name,                                                                                                                                        " & _
                                   "last_update_date,                                                                                                                                          " & _
                                   "last_updated_by,                                                                                                                                           " & _
                                   "creation_date,                                                                                                                                             " & _
                                   "created_by,                                                                                                                                                " & _
                                   "description,                                                                                                                                               " & _
                                   "query,                                                                                                                                                     " & _
                                   "parametros)                                                                                                                                                " & _
                                   "Values                                                                                                                                                     " & _
                                   "('USUARIOS_ACTIVOS',                                                                                                                                       " & _
                                   "GETDATE(),                                                                                                                                                 " & _
                                   "1,                                                                                                                                                         " & _
                                   "GETDATE(),                                                                                                                                                 " & _
                                   "1,                                                                                                                                                         " & _
                                   "'Listado de usuarios Activos',                                                                                                                             " & _
                                   "'SELECT t1.user_name as usuario, t3.description as responsabilidad                                                                                         " & _
                                   "   FROM fnd_user t1, fnd_user_resp_groups_direct t2, fnd_responsibility t3                                                                                 " & _
                                   "  WHERE t1.end_date is null and t1.user_id = t2.user_id and t2.end_date is null and t2.responsibility_id = t3.responsibility_id and t3.end_date is null;', " & _
                                   "0);                                                                                                                                                        "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'fnd_cajas
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                   " & _
                            "FROM information_schema.tables             " & _
                            "WHERE table_catalog = 'erp'          " & _
                            "AND table_name = 'fnd_cajas';    "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE fnd_cajas(                  " & _
                                   " caja_id     int         IDENTITY(1,1),  " & _
                                   " description varchar(80) NOT NULL);      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO fnd_cajas    " & _
                                   " (description)           " & _
                                   "Values                   " & _
                                   " ('Caja1');              "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO fnd_cajas    " & _
                                   " (description)           " & _
                                   "Values                   " & _
                                   " ('Caja2');              "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO fnd_cajas    " & _
                                   " (description)           " & _
                                   "Values                   " & _
                                   " ('Caja3');              "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
        '======================================================================================
        'INVENTARIOS
        '======================================================================================
            
            '..................................................................................
            'mtl_item_categories
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                   " & _
                            "FROM information_schema.tables             " & _
                            "WHERE table_catalog = 'erp'                " & _
                            "AND table_name = 'mtl_item_categories';    "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_item_categories(                                       " & _
                                   "     category_id         INT             IDENTITY(1,1)    PRIMARY KEY,  " & _
                                   "     description         VARCHAR(240),                                  " & _
                                   "     last_update_date    DATE            NOT NULL,                      " & _
                                   "     last_updated_by     INT             NOT NULL,                      " & _
                                   "     creation_date       DATE            NOT NULL,                      " & _
                                   "     created_by          INT             NOT NULL,                      " & _
                                   "     last_update_login   INT                                            " & _
                                   ");                                                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Riego',              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Seguridad',          " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipos de Ventilacion',       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Ferreteria',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Fertilizantes',                " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Fletes',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Ganado',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Gas LP',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'General',                      " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Guias',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Herramienta',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Herreria',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Higiene',                      " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Hilos',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Honorarios',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Implemento',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Abarrotes',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Acero',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Alimentacion',                 " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Aminoacidos',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Ampliacion',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Antibiotico',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Anticipos',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Aves',                         " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Biologicos',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Bovino',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Campo',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Cerdos',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Cintas',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Combustibles y Lubricantes',   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Complemento',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Concentrado',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Conductores',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Conexion',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Construccion',                 " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Control de Roedores',          " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Cuidado Animal',               " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Cuotas y Subscripciones',      " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Desinfectante',                " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Electrico',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Empaque',                      " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Computo',            " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Comunicacion',       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Granja',             " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Laboratorio',        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Oficina',            " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Equipo de Planta',             " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Importaciones',                " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Impuestos',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Insecticida',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Investigacion y Desarrollo',   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Limpieza',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Madera',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Mantenimiento y Reparacion',   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Maquinaria y Equipo',          " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Materia Prima',                " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Medicamentos',                 " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Minerales',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Mobiliario',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Muebles y Enseres',            " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'No aplica',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Nomina',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Nutricion',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Papelerias',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Paqueteria',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Pastas',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Pegamentos',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Pintura',                      " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Plastico',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Produccion',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Producto Terminado',           " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Quimicos',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Refacciones',                  " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Refrigerados',                 " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Semillas',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Servicio',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Soldaduras',                   " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Unguento',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Uniforme',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Varios',                       " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Vehiculos',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Viaticos',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_item_categories(    " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Vitaminas',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'mtl_units_of_measure
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                   " & _
                            "FROM information_schema.tables             " & _
                            "WHERE table_catalog = 'erp'                " & _
                            "AND table_name = 'mtl_units_of_measure';   "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_units_of_measure(                                      " & _
                                   "    row_id               INT            IDENTITY(1,1)   PRIMARY KEY,    " & _
                                   "    unit_of_measure      VARCHAR(25)    NOT NULL,                       " & _
                                   "    uom_code             VARCHAR(3)     NOT NULL,                       " & _
                                   "    uom_class            VARCHAR(10)    NOT NULL,                       " & _
                                   "    last_update_date     DATE           NOT NULL,                       " & _
                                   "    last_updated_by      INT            NOT NULL,                       " & _
                                   "    creation_date        DATE           NOT NULL,                       " & _
                                   "    created_by           INT            NOT NULL,                       " & _
                                   "    last_update_login    INT,                                           " & _
                                   "    description          VARCHAR(50),                                   " & _
                                   "    attribute_category   VARCHAR(30),                                   " & _
                                   "    attribute1           VARCHAR(150),                                  " & _
                                   "    attribute2           VARCHAR(150),                                  " & _
                                   "    attribute3           VARCHAR(150),                                  " & _
                                   "    attribute4           VARCHAR(150),                                  " & _
                                   "    attribute5           VARCHAR(150),                                  " & _
                                   "    attribute6           VARCHAR(150),                                  " & _
                                   "    attribute7           VARCHAR(150),                                  " & _
                                   "    attribute8           VARCHAR(150),                                  " & _
                                   "    attribute9           VARCHAR(150),                                  " & _
                                   "    attribute10          VARCHAR(150)                                   " & _
                                   ");                                                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Pieza',                        " & _
                                   "    'PZ',                           " & _
                                   "    'Cantidad',                     " & _
                                   "    'Pieza',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Watt',                         " & _
                                   "    'WT',                           " & _
                                   "    'Electrica',                    " & _
                                   "    'Watt',                         " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Metro',                        " & _
                                   "    'MT',                           " & _
                                   "    'Longitud',                     " & _
                                   "    'Metro',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Kilogramo',                    " & _
                                   "    'KG',                           " & _
                                   "    'Peso',                         " & _
                                   "    'Kilogramo',                    " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Servicio',                     " & _
                                   "    'SR',                           " & _
                                   "    'Servicio',                     " & _
                                   "    'Servicio',                     " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_units_of_measure(   " & _
                                   "    unit_of_measure,                " & _
                                   "    uom_code,                       " & _
                                   "    uom_class,                      " & _
                                   "    description,                    " & _
                                   "    last_update_date,               " & _
                                   "    last_updated_by,                " & _
                                   "    creation_date,                  " & _
                                   "    created_by,                     " & _
                                   "    last_update_login               " & _
                                   ")Values(                            " & _
                                   "    'Litro',                        " & _
                                   "    'LT',                           " & _
                                   "    'Volumen',                      " & _
                                   "    'Litro',                        " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    GETDATE(),                      " & _
                                   "    1,                              " & _
                                   "    1                               " & _
                                   ");                                  "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'mtl_transaction_types
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                   " & _
                            "FROM information_schema.tables             " & _
                            "WHERE table_catalog = 'erp'                " & _
                            "AND table_name = 'mtl_transaction_types';  "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_transaction_types(                                         " & _
                                   "     transaction_type_id    INT             IDENTITY(1,1)   PRIMARY KEY,    " & _
                                   "     transaction_type_name  VARCHAR(240)    NOT NULL,                       " & _
                                   "     description            VARCHAR(240)    NOT NULL,                       " & _
                                   "     transaction_action_id  INT             NOT NULL,                       " & _
                                   "     attribute_category     VARCHAR(30),                                    " & _
                                   "     attribute1             VARCHAR(150),                                   " & _
                                   "     attribute2             VARCHAR(150),                                   " & _
                                   "     attribute3             VARCHAR(150),                                   " & _
                                   "     attribute4             VARCHAR(150),                                   " & _
                                   "     attribute5             VARCHAR(150),                                   " & _
                                   "     attribute6             VARCHAR(150),                                   " & _
                                   "     attribute7             VARCHAR(150),                                   " & _
                                   "     attribute8             VARCHAR(150),                                   " & _
                                   "     attribute9             VARCHAR(150),                                   " & _
                                   "     attribute10            VARCHAR(150),                                   " & _
                                   "     last_update_date       DATE            NOT NULL,                       " & _
                                   "     last_updated_by        INT             NOT NULL,                       " & _
                                   "     creation_date          DATE            NOT NULL,                       " & _
                                   "     created_by             INT             NOT NULL,                       " & _
                                   "     last_update_login      INT                                             " & _
                                   ");                                                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Inventory issue',                  " & _
                                   "    'Issue material against account',   " & _
                                   "    2,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Inventory receipt',                " & _
                                   "    'Receive material against account', " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Sales order issue',                " & _
                                   "    'Ship Confirm external Sales Order'," & _
                                   "    2,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'WIP Issue',                        " & _
                                   "    'WIP Issue',                        " & _
                                   "    2,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'WIP Return',                       " & _
                                   "    'WIP Return',                       " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'WIP Completion Return',            " & _
                                   "    'WIP Completion Return',            " & _
                                   "    2,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'WIP Completion',                   " & _
                                   "    'WIP Completion',                   " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'PO Receipt',                       " & _
                                   "    'Receive Purchase Order',           " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(      " & _
                                   "     transaction_type_name,             " & _
                                   "     description,                       " & _
                                   "     transaction_action_id,             " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Return to Vendor',                 " & _
                                   "    'Return to vendor from stores',     " & _
                                   "    2,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(                          " & _
                                   "     transaction_type_name,                                 " & _
                                   "     description,                                           " & _
                                   "     transaction_action_id,                                 " & _
                                   "     last_update_date,                                      " & _
                                   "     last_updated_by,                                       " & _
                                   "     creation_date,                                         " & _
                                   "     created_by,                                            " & _
                                   "     last_update_login                                      " & _
                                   ")Values(                                                    " & _
                                   "    'Positive Physical Inv Adjust',                         " & _
                                   "    'Positive Physical Inventory adjustment transactions',  " & _
                                   "    1,                                                      " & _
                                   "    GETDATE(),                                              " & _
                                   "    1,                                                      " & _
                                   "    GETDATE(),                                              " & _
                                   "    1,                                                      " & _
                                   "    1                                                       " & _
                                   ");                                                          "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_transaction_types(                          " & _
                                   "     transaction_type_name,                                 " & _
                                   "     description,                                           " & _
                                   "     transaction_action_id,                                 " & _
                                   "     last_update_date,                                      " & _
                                   "     last_updated_by,                                       " & _
                                   "     creation_date,                                         " & _
                                   "     created_by,                                            " & _
                                   "     last_update_login                                      " & _
                                   ")Values(                                                    " & _
                                   "    'Negative Physical Inv Adjust',                         " & _
                                   "    'Negative Physical Inventory adjustment transactions',  " & _
                                   "    2,                                                      " & _
                                   "    GETDATE(),                                              " & _
                                   "    1,                                                      " & _
                                   "    GETDATE(),                                              " & _
                                   "    1,                                                      " & _
                                   "    1                                                       " & _
                                   ");                                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
    
            '..................................................................................
            'mtl_material_transactions
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                       " & _
                            "FROM information_schema.tables                 " & _
                            "WHERE table_catalog = 'erp'                    " & _
                            "AND table_name = 'mtl_material_transactions';  "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_material_transactions(                                                                                         " & _
                                   "     transaction_id         INT             IDENTITY(1,1)   PRIMARY KEY,    inventory_item_id      INT             NOT NULL,    " & _
                                   "     subinventory_code      VARCHAR(240)    NOT NULL,                       transaction_type_id    INT             NOT NULL,    " & _
                                   "     transaction_action_id  INT             NOT NULL,                       transaction_source     VARCHAR(150)    NOT NULL,    " & _
                                   "     uom                    VARCHAR(240)    NOT NULL,                       transaction_quantity   INT             NOT NULL,    " & _
                                   "     actual_cost            INT             NOT NULL,                       transaction_cost       INT             NOT NULL,    " & _
                                   "     lot_number             VARCHAR(240),                                   attribute_category     VARCHAR(30),                 " & _
                                   "     attribute1             VARCHAR(150),                                   attribute2             VARCHAR(150),                " & _
                                   "     attribute3             VARCHAR(150),                                   attribute4             VARCHAR(150),                " & _
                                   "     attribute5             VARCHAR(150),                                   attribute6             VARCHAR(150),                " & _
                                   "     attribute7             VARCHAR(150),                                   attribute8             VARCHAR(150),                " & _
                                   "     attribute9             VARCHAR(150),                                   attribute10            VARCHAR(150),                " & _
                                   "     attribute11            VARCHAR(150),                                   attribute12            VARCHAR(150),                " & _
                                   "     attribute13            VARCHAR(150),                                   attribute14            VARCHAR(150),                " & _
                                   "     attribute15            VARCHAR(150),                                   attribute16            VARCHAR(150),                " & _
                                   "     attribute17            VARCHAR(150),                                   attribute18            VARCHAR(150),                " & _
                                   "     attribute19            VARCHAR(150),                                   attribute20            VARCHAR(150),                " & _
                                   "     attribute21            VARCHAR(150),                                   attribute22            VARCHAR(150),                " & _
                                   "     attribute23            VARCHAR(150),                                   attribute24            VARCHAR(150),                " & _
                                   "     attribute25            VARCHAR(150),                                   attribute26            VARCHAR(150),                " & _
                                   "     attribute27            VARCHAR(150),                                   attribute28            VARCHAR(150),                " & _
                                   "     attribute29            VARCHAR(150),                                   attribute30            VARCHAR(150),                " & _
                                   "     last_update_date       DATE            NOT NULL,                       last_updated_by        INT              NOT NULL,   " & _
                                   "     creation_date          DATE            NOT NULL,                       created_by             INT              NOT NULL,   " & _
                                   "     last_update_login      INT);                                                                                               "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'mtl_system_items_b
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla               " & _
                            "FROM information_schema.tables         " & _
                            "WHERE table_catalog = 'erp'            " & _
                            "AND table_name = 'mtl_system_items_b'; "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_system_items_b(                                                                                                " & _
                                   "     inventory_item_id      INT             IDENTITY(1,1)   PRIMARY KEY,    description            VARCHAR(240)    NOT NULL,    " & _
                                   "     segment1               VARCHAR(240)    NOT NULL,                       uom                    VARCHAR(240)    NOT NULL,    " & _
                                   "     item_cost              DECIMAL         NOT NULL,                       lot_control            INT,                         " & _
                                   "     category_id            INT,                                                                                                " & _
                                   "     tax_rate               DECIMAL         NOT NULL,                       attribute_category     VARCHAR(30),                 " & _
                                   "     attribute1             VARCHAR(150),                                   attribute2             VARCHAR(150),                " & _
                                   "     attribute3             VARCHAR(150),                                   attribute4             VARCHAR(150),                " & _
                                   "     attribute5             VARCHAR(150),                                   attribute6             VARCHAR(150),                " & _
                                   "     attribute7             VARCHAR(150),                                   attribute8             VARCHAR(150),                " & _
                                   "     attribute9             VARCHAR(150),                                   attribute10            VARCHAR(150),                " & _
                                   "     attribute11            VARCHAR(150),                                   attribute12            VARCHAR(150),                " & _
                                   "     attribute13            VARCHAR(150),                                   attribute14            VARCHAR(150),                " & _
                                   "     attribute15            VARCHAR(150),                                   attribute16            VARCHAR(150),                " & _
                                   "     attribute17            VARCHAR(150),                                   attribute18            VARCHAR(150),                " & _
                                   "     attribute19            VARCHAR(150),                                   attribute20            VARCHAR(150),                " & _
                                   "     attribute21            VARCHAR(150),                                   attribute22            VARCHAR(150),                " & _
                                   "     attribute23            VARCHAR(150),                                   attribute24            VARCHAR(150),                " & _
                                   "     attribute25            VARCHAR(150),                                   attribute26            VARCHAR(150),                " & _
                                   "     attribute27            VARCHAR(150),                                   attribute28            VARCHAR(150),                " & _
                                   "     attribute29            VARCHAR(150),                                   attribute30            VARCHAR(150),                " & _
                                   "     last_update_date       DATE            NOT NULL,                       last_updated_by        INT              NOT NULL,   " & _
                                   "     creation_date          DATE            NOT NULL,                       created_by             INT              NOT NULL,   " & _
                                   "     last_update_login      INT);                                                                                               "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
            
            '..................................................................................
            'mtl_secondary_inventories
            '..................................................................................
            StTableExists = "SELECT COUNT(*) as tabla                       " & _
                            "FROM information_schema.tables                 " & _
                            "WHERE table_catalog = 'erp'                    " & _
                            "AND table_name = 'mtl_secondary_inventories';  "
            With TableExists
                If .State = 1 Then .Close
                    .Open StTableExists, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            If TableExists.Fields("tabla") = 0 Then
                With CreateTable
                    .CommandText = "CREATE TABLE mtl_secondary_inventories(                             " & _
                                   "     secondary_inventory_name   VARCHAR(20) PRIMARY KEY NOT NULL,   " & _
                                   "     description                VARCHAR(50),                        " & _
                                   "     disable_date               DATE,                               " & _
                                   "     last_update_date           DATE        NOT NULL,               " & _
                                   "     last_updated_by            INT         NOT NULL,               " & _
                                   "     creation_date              DATE        NOT NULL,               " & _
                                   "     created_by                 INT         NOT NULL,               " & _
                                   "     last_update_login          INT                                 " & _
                                   ");                                                                  "
                    .ActiveConnection = Cn
                    .Execute
                End With
                With InsertIntoTable
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Combustible',                      " & _
                                   "    'Alamacen de Combustible',          " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Comunicacion',                     " & _
                                   "    'Almacen de Comunicaciones',        " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Equipo de oficina',                " & _
                                   "    'Activo Fijo Equipo de Oficinas',   " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Herramienta',                      " & _
                                   "    'Almacen de Herramientas',          " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Insumos',                          " & _
                                   "    'Almacen de Insumos',               " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Limpieza',                         " & _
                                   "    'Almacen de Limpieza',              " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Maquinaria y equipo',              " & _
                                   "    'Activo Fijo Maquinaria y Equipo',  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Material Electrico',               " & _
                                   "    'Almacen de Material Electrico',    " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Materia Prima',                    " & _
                                   "    'Almacen de Materia Prima',         " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Papeleria',                        " & _
                                   "    'Almacen de Papeleria',             " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Pintura',                          " & _
                                   "    'Almacen de Pintura',               " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Plagas',                           " & _
                                   "    'Almacen de Control de Plagas',     " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Producto Terminado',               " & _
                                   "    'Almacen de Producto terminado',    " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Refaccion',                        " & _
                                   "    'Almacen de Refacciones',           " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Uniformes',                        " & _
                                   "    'Almacen de Uniformes',             " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                    .CommandText = "INSERT INTO mtl_secondary_inventories(  " & _
                                   "     secondary_inventory_name,          " & _
                                   "     description,                       " & _
                                   "     last_update_date,                  " & _
                                   "     last_updated_by,                   " & _
                                   "     creation_date,                     " & _
                                   "     created_by,                        " & _
                                   "     last_update_login                  " & _
                                   ")Values(                                " & _
                                   "    'Vehiculos',                        " & _
                                   "    'Activo Fijo Vehiculos en Servicio'," & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    GETDATE(),                          " & _
                                   "    1,                                  " & _
                                   "    1                                   " & _
                                   ");                                      "
                    .ActiveConnection = Cn
                    .Execute
                End With
                TableExists.Close
            End If
End Sub
Sub CloseBd()
    'Cerramos la conexion a la base de datos
    With Cn
        If .State = 1 Then
            .Close
        End If
    End With
End Sub




'==============================================================================================

'F  O   R   M   U   L   A   R   I   O   S

'==============================================================================================
Sub main()
    'Abrimos la conexion a la base de datos
    OpenBd
    Form1.Show
End Sub
Sub ProgramExit()
    'cerramos el programa
    ClearFrames
    CloseBd
    Unload Form1
End Sub
Sub ClearFrames()
    Dim i As Integer
    With Form1
        .Left = 0
        .Top = 0
        '--------------------------------------------------------------------------------------
        'Menu Buscar
        '--------------------------------------------------------------------------------------
        .FBuscar.Visible = False
        .txtBuscar.Text = ""
        .lstBuscar.Clear
        
        '--------------------------------------------------------------------------------------
        'Menu Acciones
        '--------------------------------------------------------------------------------------
        .CNuevo.Enabled = False
        .CBuscar.Enabled = False
        .CGuardar.Enabled = False
        .CImprimir.Enabled = False
        .CEliminar.Enabled = False
        
        '--------------------------------------------------------------------------------------
        'Menu Atributos
        '--------------------------------------------------------------------------------------
        For i = 1 To 30
            .txtAtributos(i).Text = ""
        Next i
        
        '--------------------------------------------------------------------------------------
        'ADMINISTRADOR DEL SISTEMA
        '--------------------------------------------------------------------------------------
            '==================================================================================
            'Grupos de concurrentes
            '==================================================================================
            .FGruposConcurrentes.Visible = False
            .TGruposConcurrentes.Text = ""
            .LGruposConcurrentes.Clear
            
            '==================================================================================
            'Cajas
            '==================================================================================
            .FCajas.Visible = False
            .txtCajas.Text = ""
            .lstCajas.Clear
            
            '==================================================================================
            'Usuarios
            '==================================================================================
            .FUsuarios.Visible = False
            For i = 0 To 2
                .txtusuarios(i).Text = ""
                .txtusuarios(i).Visible = True
                .txtAltaUsuarios(i).Text = ""
                .txtAltaUsuarios(i).Visible = False
            Next i
            .lstusuarios.Clear
            
            '==================================================================================
            'Creacion de Reportes
            '==================================================================================
            .FCreacionReportes.Visible = False
            For i = 0 To 2
                .txtCreacionReportes(i).Text = ""
                .txtCreacionReportes(i).Visible = True
                .txtNCreacionReportes(i).Text = ""
                .txtNCreacionReportes(i).Visible = False
            Next i
            .cbCreacionReportes.Value = 0
            .cbCreacionReportes.Visible = True
            .cbNCreacionReportes.Value = 0
            .cbNCreacionReportes.Visible = False
            
            '==================================================================================
            'Ejecucion de Reportes
            '==================================================================================
            .FEjecutarReporte.Visible = False
            .txtEjecutarReporte.Text = ""
            For i = 0 To 1
                .DTPEjecutarReporte(i).Value = Date
            Next i
            .CDReportes.FileName = ""
            
        '--------------------------------------------------------------------------------------
        'INVENTARIOS
        '--------------------------------------------------------------------------------------
            '==================================================================================
            'Articulos
            '==================================================================================
            For i = 0 To 5
                .txtArticulo(i).Text = ""
                .txtArticulo(i).Visible = True
                .txtNArticulo(i).Text = ""
                .txtNArticulo(i).Visible = False
            Next i
            For i = 0 To 1
                .CbtArticulos(i).Visible = False
            Next i
            .cbArticulos.Value = 0
            .cbArticulos.Visible = True
            .cbNArticulos.Value = 0
            .cbNArticulos.Visible = False
    End With
    With SubMenuRs1
        If .State = 1 Then .Close
    End With
End Sub
