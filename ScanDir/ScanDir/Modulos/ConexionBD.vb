Imports System.Data
Imports System.Data.Common
Imports System.Data.SQLite
Imports System.Windows.Forms

Module ConexionBD
    Public ProData As DbProviderFactory
    Public MainConex As SQLiteConnection
    Dim dA As DbDataAdapter
    Dim cmdSQL As DbCommand
    Enum TiposBase
        SQLServer = 1
        Oracle = 2
        MySQL = 3
        Access2003 = 4
        Access2007 = 5
        PostgreSQL = 6
        SQLlite = 7
    End Enum

    ''' <summary>
    ''' Funcion para conectar mediante ADO.NET a bases de datos
    ''' </summary>
    ''' <param name="Tipo">Tipo de la base de datos</param>
    ''' <param name="IPServer">IP del Servidor de Bases de Datos</param>
    ''' <param name="Puerto">Puerto de Conexión a la base de datos</param>
    ''' <param name="Usuario">Usuario de acceso a la base de datos</param>
    ''' <param name="Passw">Password del Usuario</param>
    ''' <param name="Servicio">Servicio (Oracle), Catalogo (SQLServer) o Instancia (MySQL)</param>
    ''' <param name="RutaFile">Ruta del fichero Access o SQLlite</param>
    ''' <returns>Devuelve un booleano indicando si hay o no conexion y crea un objeto conexion MainConex</returns>
    ''' <remarks></remarks>
    Function ConectarBD(ByVal Tipo As TiposBase, ByVal IPServer As String, ByVal Puerto As Long, ByVal Usuario As String, _
                       ByVal Passw As String, ByVal Servicio As String, ByVal RutaFile As String) As Boolean

        Dim sProveedor As String
        Dim CadenaConexion As String

        ConectarBD = False
        CadenaConexion = ""
        sProveedor = ""

        Try
            MainConex = New SQLiteConnection()
            MainConex.ConnectionString = "Data Source=" & RutaFile & ";Version=3;"
            MainConex.Open()
            ConectarBD = True
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        Return True


        If Tipo = 1 Then
            sProveedor = "System.Data.SqlClient"
            CadenaConexion = "server=" & IPServer & "," & Puerto & ";" & _
                             "uid=" & Usuario & ";pwd=" & Passw & ";Initial Catalog=" & Servicio
        ElseIf Tipo = 2 Then
            sProveedor = "System.Data.OracleClient"
            CadenaConexion = "User Id=" & Usuario & "; Password=" & Passw & ";" & _
                            "Data Source=(DESCRIPTION = (" & _
                            "ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)" & _
                            "(HOST = " & IPServer & ")(PORT = " & Puerto & ")) )" & _
                            "(CONNECT_DATA = (SERVER = DEDICATED) " & _
                            "(SERVICE_NAME = " & Servicio & ")));"
        ElseIf Tipo = 3 Then
            sProveedor = "MySql.Data.MySqlClient"
            CadenaConexion = "server=" & IPServer & ";user id=" & Usuario & ";port=" & Puerto & ";" & _
                                "password=" & Passw & ";database=" & Servicio & ";pooling=false"
        ElseIf Tipo = 4 Then
            sProveedor = "OleDb.OleDbConnection"
            If Usuario = "" And Passw = "" Then
                CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & RutaFile
            ElseIf Usuario = "" And Passw <> "" Then
                CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaFile & ";" & _
                                    "Jet OLEDB:Database Password=" & Passw & ";"
            ElseIf Usuario <> "" And Passw <> "" Then
                CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                "Data Source=" & RutaFile & ";" & _
                                "Jet OLEDB:System Database=system.mdw;" & _
                                "User ID=" & Usuario & ";Password=" & Passw & ";"
            End If
        ElseIf Tipo = 5 Then
            sProveedor = "OleDb.OleDbConnection"
            If Usuario = "" And Passw = "" Then
                CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & RutaFile & ";Persist Security Info=False"
            ElseIf Usuario = "" And Passw <> "" Then
                CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source" & RutaFile & ";Jet OLEDB:Database Password=" & Passw & ";"
            End If
        ElseIf Tipo = 6 Then
            sProveedor = "Npgsql"
            CadenaConexion = "Server=" & IPServer & ";" & _
                            "Port=" & Puerto & ";" & _
                            "User Id=" & Usuario & ";" & _
                            "Password=" & Passw & ";" & _
                            "Database=" & Servicio & ";"
        ElseIf Tipo = 7 Then
            sProveedor = "System.Data.SQLite"
            CadenaConexion = "Data Source=" & RutaFile & ";Pooling=true;FailIfMissing=false"
        Else
            Exit Function
        End If

        ProData = DbProviderFactories.GetFactory(sProveedor)
        Try
            MainConex = ProData.CreateConnection
            MainConex.ConnectionString = CadenaConexion
            MainConex.Open()
            ConectarBD = True
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Function


    Function DesconectarBD() As Boolean

        If MainConex.State = ConnectionState.Closed Then Exit Function
        Try
            MainConex.Close()
            MainConex.Dispose()
            DesconectarBD = True
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message)
        End Try
    End Function

 
    ''' <summary>
    ''' Funcion para crear un recordset utilizando el modelo de objetos ADO.NET
    ''' </summary>
    ''' <param name="CadenaSQL">Cadena SQL a ejecutar</param>
    ''' <param name="Contenedor">Objeto Datatable para guardar la información</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CargarRecordset(ByVal CadenaSQL As String, ByRef Contenedor As DataTable) As Boolean

        cmdSQL = ProData.CreateCommand()
        cmdSQL.Connection = MainConex
        cmdSQL.CommandType = CommandType.Text
        cmdSQL.CommandText = CadenaSQL
        dA = ProData.CreateDataAdapter()
        'Contenedor = New DataTable
        dA.SelectCommand = cmdSQL
        Try
            dA.Fill(Contenedor)
            CargarRecordset = True
        Catch Fallo As Exception
            CargarRecordset = False
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            dA.Dispose()
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            cmdSQL.Dispose()
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    ''' <summary>
    ''' Crea una transacción y ejecuta una serie de sentencias
    ''' </summary>
    ''' <param name="ListaSQL">Sentencias a procesar</param>
    ''' <returns></returns>
    Function ExeTran(ByRef ListaSQL() As String) As Boolean

        Dim myTrans As SQLiteTransaction
        Dim myComando As SQLiteCommand
        Dim SQLfrase As String
        ExeTran = False

        myTrans = MainConex.BeginTransaction
        Try
            myComando = New SQLiteCommand
            myComando.Connection = MainConex
            myComando.CommandType = CommandType.Text
            For Each SQLfrase In ListaSQL
                'myComando.CommandText = "UPDATE DOCSIDDAE SET Tipo='Actas de Deslindes' WHERE iddoc=190"
                If SQLfrase = "" Then Continue For
                myComando.CommandText = SQLfrase
                myComando.Transaction = myTrans
                myComando.ExecuteNonQuery()
            Next
            myComando.Dispose()
            myTrans.Commit()
            ExeTran = True
        Catch ex As Exception
            myTrans.Rollback()
            MessageBox.Show(ex.Message)
        Finally
            myTrans.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' Crea una transacción y ejecuta una serie de sentencias
    ''' </summary>
    ''' <param name="ListaSQL">Sentencias a procesar</param>
    ''' <returns></returns>
    Function ExeTran(ByRef ListaSQL As ArrayList) As Boolean

        Dim myTrans As SQLiteTransaction
        Dim myComando As SQLiteCommand
        Dim SQLfrase As String
        ExeTran = False
        myTrans = MainConex.BeginTransaction
        Try
            myComando = New SQLiteCommand
            myComando.Connection = MainConex
            myComando.CommandType = CommandType.Text
            For Each SQLfrase In ListaSQL
                If SQLfrase = "" Then Continue For
                myComando.CommandText = SQLfrase
                myComando.Transaction = myTrans
                myComando.ExecuteNonQuery()
            Next
            myComando.Dispose()
            myTrans.Commit()
            ExeTran = True
        Catch ex As Exception
            myTrans.Rollback()
            MessageBox.Show(ex.Message)
        Finally
            myTrans.Dispose()
        End Try

    End Function



    Function CargarDatatable(ByVal CadenaSQL As String, ByRef Contenedor As DataTable) As Boolean

        Dim dataR As DbDataReader


        cmdSQL = ProData.CreateCommand()
        cmdSQL.Connection = MainConex
        cmdSQL.CommandType = CommandType.Text
        cmdSQL.CommandText = CadenaSQL
        Try
            dataR = cmdSQL.ExecuteReader()
            Contenedor.Load(dataR)
            dataR.Close()
            dataR = Nothing
            cmdSQL.Dispose()
            Application.DoEvents()
            CargarDatatable = True
        Catch Fallo As Exception
            CargarDatatable = False
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    ''' <summary>
    ''' Ejecuta una sentenia sin Transacciones
    ''' </summary>
    ''' <param name="cadSQL">Sentencia SQL</param>
    ''' <returns></returns>
    Function ExeSinTran(ByVal cadSQL As String) As Boolean

        Dim myComando As SQLiteCommand

        If cadSQL = "" Then Exit Function
        ExeSinTran = False
        Try
            myComando = New SQLiteCommand
            myComando.Connection = MainConex
            myComando.CommandType = CommandType.Text
            myComando.CommandText = cadSQL
            myComando.ExecuteNonQuery()
            myComando.Dispose()
            ExeSinTran = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, AplicacionTitulo, MessageBoxButtons.OK, MessageBoxIcon.Error)
            GenerarLOG(cadSQL)
        End Try


    End Function

    ''' <summary>
    ''' Devuelve el resultado único de una SQL almacenada en un integer
    ''' </summary>
    ''' <param name="cadenaSQL">Sentencia SQL</param>
    ''' <param name="valCommand">Valor por referencia</param>
    ''' <returns></returns>
    Function ObtenerEscalarNumerico(ByVal cadenaSQL As String, ByRef valCommand As Integer) As Boolean

        If MainConex.State <> ConnectionState.Open Then
            MessageBox.Show("La conexión con DB está cerrada", AplicacionTitulo, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        cmdSQL = New SQLiteCommand

        Try
            cmdSQL.Connection = MainConex
            cmdSQL.CommandType = CommandType.Text
            cmdSQL.CommandText = cadenaSQL
            valCommand = Convert.ToInt32(cmdSQL.ExecuteScalar())
            Return True
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True
        Finally
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End Try



    End Function

    ''' <summary>
    ''' Devuelve el resultado único de una SQL almacenada en un string
    ''' </summary>
    ''' <param name="cadenaSQL">Sentencia SQL</param>
    ''' <param name="textCommand">Valor por referencia</param>
    ''' <returns></returns>
    Function ObtenerEscalarTexto(ByVal cadenaSQL As String, ByRef textCommand As String) As Boolean

        If MainConex.State <> ConnectionState.Open Then
            MessageBox.Show("La conexión con DB está cerrada", AplicacionTitulo, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        cmdSQL = New SQLiteCommand

        Try
            cmdSQL.Connection = MainConex
            cmdSQL.CommandType = CommandType.Text
            cmdSQL.CommandText = cadenaSQL
            textCommand = Convert.ToString(cmdSQL.ExecuteScalar())
            Return True
        Catch Fallo As Exception
            MessageBox.Show(Fallo.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True
        Finally
            cmdSQL.Dispose()
            cmdSQL = Nothing
        End Try



    End Function

End Module
