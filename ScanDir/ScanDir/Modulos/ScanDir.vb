Imports System.Windows.Forms

Public Class ScanDir
    Public NumeroFicheros As String
    Public NumeroDirectorios As String
    Public DIRBase As String
    Public ProcesarJPGcabeceras As Boolean
    Dim BasePreparada As Boolean
    Dim cadCommand As String

    Property AnalizarJPGcabeceras() As Boolean
        Get
            Return ProcesarJPGcabeceras
        End Get
        Set(ByVal value As Boolean)
            ProcesarJPGcabeceras = value
        End Set
    End Property
    Function ListarFicherosDIR(ByVal DirectorioIN As String,
                                Optional ByRef Extensiones() As String = Nothing) As Long

        Dim directoryPaths As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim directorySubPaths As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        'Dim directoryPaths() As String
        Dim directoryPath As String
        Dim directorySubPath As String
        Dim NombreBaseFinal As String
        Dim databaseDir As String
        databaseDir = My.Computer.FileSystem.SpecialDirectories.Desktop & "\dirdisk.db3"
        'databaseDir = System.Environment.SpecialFolder.DesktopDirectory & "\dirdisk.db3"
        If BasePreparada = False Then
            'No se puede localizar la base de datos de trabajo
            ListarFicherosDIR = 99
            Exit Function
        End If
        If My.Computer.FileSystem.FileExists(databaseDir) = True Then
            System.IO.File.Delete(databaseDir)
        End If
        If ConectarBD(TiposBase.SQLlite, "", 0, "", "", "", databaseDir) = False Then
            'No se puede conectar a la base de datos
            ListarFicherosDIR = 98
            Exit Function
        End If
        'Creamos la base de datos

        cadCommand = "CREATE TABLE rutasfiles (" &
                    "ruta varchar(512)," &
                    "nombrefich varchar(512)," &
                    "nombredir varchar(512)," &
                    "numbytes integer," &
                    "fecha_creacion date," &
                    "fecha_ultimaescritura date," &
                    "fecha_ultimoacceso date," &
                    "extension varchar(64)," &
                    "pixelwidth integer," &
                    "pixelheight integer," &
                    "reswidth integer," &
                    "resheight integer" &
                    ")"
        ExeSinTran(cadCommand)

        'Creamos los objetos comando
        'Iniciamos las variables
        NumeroFicheros = 0
        NumeroDirectorios = 0
        DIRBase = ""
        'Sacamos los ficheros del directorio de entrada
        SacarListaArchivos(DirectorioIN, Extensiones)
        Try
            directoryPaths = My.Computer.FileSystem.GetDirectories(DirectorioIN, FileIO.SearchOption.SearchTopLevelOnly)
        Catch e As Exception
            Application.DoEvents()
        End Try

        'Recorremos los directorios de primer nivel del directorio de entrada
        For Each directoryPath In directoryPaths
            Try
                SacarListaArchivos(directoryPath, Extensiones)
                'Recorremos los directorios y subdirectorios de cada directorio de primer nivel del directorio de entrada
                directorySubPaths = My.Computer.FileSystem.GetDirectories(directoryPath, FileIO.SearchOption.SearchAllSubDirectories)
                NumeroDirectorios = NumeroDirectorios + 1
                For Each directorySubPath In directorySubPaths
                    NumeroDirectorios = NumeroDirectorios + 1
                    SacarListaArchivos(directorySubPath, Extensiones)
                Next
            Catch ex As Exception
                Application.DoEvents()
            End Try
        Next
        'Desconectamos la base de datos
        DesconectarBD()

        'If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\dirdisk.db3") = True Then
        '    NombreBaseFinal = My.Application.Info.DirectoryPath & "\Scandir_" & Replace(Replace(CStr(Now), "/", "-"), ":", "_") & ".db3"
        '    System.IO.File.Move(My.Application.Info.DirectoryPath & "\dirdisk.db3", NombreBaseFinal)
        '    DIRBase = NombreBaseFinal
        'End If


        ListarFicherosDIR = 0

    End Function

    Sub SacarListaArchivos(ByVal Directorio As String, Optional ByRef Extensiones() As String = Nothing)


        Dim NombreFicheros As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim NombreFichero As String
        Dim DatosFichero As System.IO.FileInfo
        Dim CargarEnBase As Boolean
        'Variables para guardar los parametros del JPG
        Dim Anchura As Integer
        Dim Altura As Integer
        Dim ResX As Integer
        Dim ResY As Integer
        Dim ListaSQL As New ArrayList

        NombreFicheros = My.Computer.FileSystem.GetFiles(Directorio)
        For Each NombreFichero In NombreFicheros
            DatosFichero = My.Computer.FileSystem.GetFileInfo(NombreFichero)
            NumeroFicheros = NumeroFicheros + 1
            MainForm.ToolStripStatusLabel1.Text = NumeroFicheros
            Application.DoEvents()
            '-------------------------------------------------------------------------
            CargarEnBase = False
            If IsNothing(Extensiones) Then
                CargarEnBase = True
            Else
                CargarEnBase = IIf(Array.IndexOf(Extensiones, DatosFichero.Extension.ToLower) >= 0, True, False)
            End If
            If CargarEnBase = True Then
                If ProcesarJPGcabeceras = False Then
                    cadCommand = "INSERT INTO rutasfiles " &
                            "(ruta,nombrefich,nombredir,numbytes,fecha_creacion,fecha_ultimaescritura,fecha_ultimoacceso," &
                            "extension) VALUES " &
                            "(""" & DatosFichero.FullName & """," &
                            """" & DatosFichero.Name & """," &
                            """" & DatosFichero.DirectoryName & """," &
                            "" & DatosFichero.Length & "," &
                            "'" & String.Format("{0:s}", DatosFichero.CreationTime) & "'," &
                            "'" & String.Format("{0:s}", DatosFichero.LastWriteTime) & "'," &
                            "'" & String.Format("{0:s}", DatosFichero.LastAccessTime) & "'," &
                            """" & DatosFichero.Extension & """)"
                    ListaSQL.Add(cadCommand)
                Else
                    Anchura = 0
                    Altura = 0
                    ResX = 0
                    ResY = 0
                    If DatosFichero.Extension.ToLower = ".jpg" Then _
                                    AnalizarJPGHeader(DatosFichero.FullName, Anchura, Altura, ResX, ResY)
                    cadCommand = "INSERT INTO rutasfiles " &
                            "(Ruta,nombrefich,nombredir,numbytes,fecha_creacion,fecha_ultimaescritura,fecha_ultimoacceso," &
                            "extension,pixelwidth,pixelheight,reswidth,resheight) VALUES " &
                            "(""" & DatosFichero.FullName & """," &
                            """" & DatosFichero.Name & """," &
                            """" & DatosFichero.DirectoryName & """," &
                            "" & DatosFichero.Length & "," &
                            "'" & String.Format("{0:s}", DatosFichero.CreationTime) & "'," &
                            "'" & String.Format("{0:s}", DatosFichero.LastWriteTime) & "'," &
                            "'" & String.Format("{0:s}", DatosFichero.LastAccessTime) & "'," &
                            """" & DatosFichero.Extension & """," &
                            CType(Anchura, Integer) & "," &
                            CType(Altura, Integer) & "," &
                            CType(ResX, Integer) & "," &
                            CType(ResY, Integer) & "" &
                            ")"
                    ListaSQL.Add(cadCommand)
                    '------------------------------------------------------------------------------
                End If
            End If
        Next
        ExeTran(ListaSQL)

    End Sub

    Public Sub New()

        'If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\semilla.dat") = True Then
        '    Try
        '        FileCopy(My.Application.Info.DirectoryPath & "\semilla.dat", My.Application.Info.DirectoryPath & "\dirdisk.dat")
        '        BasePreparada = True
        '    Catch E As Exception
        '        MsgBox(E.Message)
        '        BasePreparada = False
        '    End Try
        'Else
        '    BasePreparada = False
        'End If

        BasePreparada = True
        ProcesarJPGcabeceras = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="RutaFichero">Ruta de entrada del fichero JPG</param>
    ''' <param name="ImageWidth">Variable para almacenar anchura en pixel.(Número de columnas)</param>
    ''' <param name="ImageHeight">Variable para almacenar altura en pixel.(Numero de Lineas)</param>
    ''' <param name="XDensidad">Densidad Eje X</param>
    ''' <param name="YDensidad">Densidad Eje Y</param>
    ''' <returns>
    ''' 0:Proceso OK
    ''' 99: El fichero no existe
    ''' 98: El fichero no es un JPG
    ''' </returns>
    ''' <remarks></remarks>
    Function AnalizarJPGHeader(ByVal RutaFichero As String,
                                    ByRef ImageWidth As Integer,
                                    ByRef ImageHeight As Integer,
                                    ByRef XDensidad As Long,
                                    ByRef YDensidad As Long) As Long

        Dim FileCad As IO.FileStream
        Dim FileBin As IO.BinaryReader
        Dim Letra As Byte
        Dim okProc As Boolean
        Dim CabJFIF As String
        Dim CabMarcador As String
        Dim CabLongitud As Integer
        Dim iSubBucle As Long
        Dim Unidades As Byte
        Dim VersionJFIF As Integer
        Dim PosicionINICab As Long
        Dim TXT As String
        'Comprobamos existencia del fichero
        If System.IO.File.Exists(RutaFichero) = False Then
            AnalizarJPGHeader = 99  'El fichero no existe
            Exit Function
        End If
        'Abrimos el fichero
        Try
            FileCad = New IO.FileStream(RutaFichero,
                                                IO.FileMode.Open,
                                                IO.FileAccess.Read,
                                                IO.FileShare.Read)
            FileBin = New IO.BinaryReader(FileCad)
        Catch Manage_err As Exception
            AnalizarJPGHeader = 98  'Error al acceder el fichero
            Exit Function
        End Try
        'Inicializamos las variables
        CabMarcador = ""
        CabJFIF = ""
        ImageWidth = 0
        ImageHeight = 0
        'SOI del JPG (Start Of Image)
        If Hex(FileBin.ReadByte) = "FF" Then okProc = True
        If Hex(FileBin.ReadByte) = "D8" Then okProc = True
        If okProc = False Then
            AnalizarJPGHeader = 97  'El fichero no es un JPG
        End If
        Do
            'Localizada una cabecera y calculo de su longitud en bytes
            If Hex(FileBin.ReadByte) = "FF" Then
                CabMarcador = Hex(FileBin.ReadByte)
                PosicionINICab = FileBin.BaseStream.Position - 1
            End If
            CabLongitud = 256 * FileBin.ReadByte
            CabLongitud = CabLongitud + FileBin.ReadByte
            'Procesamos en función de la cabecera
            Select Case CabMarcador
                Case Is = "E0"  'Marcador Application Marker APP0
                    For iSubBucle = 1 To 4
                        CabJFIF = CabJFIF & FileBin.ReadChar
                    Next iSubBucle
                    If CabJFIF = "JFIF" Then     'Identificador JFIF
                        Letra = FileBin.ReadByte
                        VersionJFIF = 256 * FileBin.ReadByte
                        VersionJFIF = VersionJFIF + FileBin.ReadByte
                        '257:Version 1.1
                        '258:Version 1.2
                        Unidades = FileBin.ReadByte
                        XDensidad = 256 * FileBin.ReadByte
                        XDensidad = XDensidad + FileBin.ReadByte
                        YDensidad = 256 * FileBin.ReadByte
                        YDensidad = YDensidad + FileBin.ReadByte
                        Application.DoEvents()
                    Else
                        AnalizarJPGHeader = 98
                        Exit Do
                    End If
                Case Is = "EC"  'Marcador APP12
                    Application.DoEvents()
                Case Is = "EE"  'Marcador APP14
                    Application.DoEvents()
                Case Is = "DB"  'Marcador DQT (Define a Quantization Table)
                    Application.DoEvents()
                Case Is = "C0"  'Marcador S0F0 (Baseline DCT)
                    Application.DoEvents()
                    Letra = FileBin.ReadByte
                    ImageHeight = 256 * FileBin.ReadByte
                    ImageHeight = ImageHeight + FileBin.ReadByte
                    ImageWidth = 256 * FileBin.ReadByte
                    ImageWidth = ImageWidth + FileBin.ReadByte
                    Application.DoEvents()
                    AnalizarJPGHeader = 0
                    Exit Do
                Case Is = "C4"  'Marcador DHT (Define Huffman Table)
                    Application.DoEvents()
                Case Is = "E1"
                    'TXT = ""
                    'For iSubBucle = 4 To CabLongitud
                    ' TXT = TXT & FileBin.ReadChar
                    'Next
                    Application.DoEvents()
                Case Is = "DA"  'Marcador SOS (Start Of Scan)
                    Application.DoEvents()
                Case Else
                    Application.DoEvents()
            End Select
            Application.DoEvents()
            CabMarcador = ""
            FileBin.BaseStream.Position = PosicionINICab + CabLongitud + 1
        Loop
        FileBin.Close()
        FileBin = Nothing
        FileCad.Close()
        FileCad.Dispose()
    End Function

    Function AnalizarJPGHeader(ByVal RutaFichero As String, ByVal container As ListBox) As Long

        Dim FileCad As IO.FileStream
        Dim FileBin As IO.BinaryReader
        Dim Letra As Byte
        Dim okProc As Boolean
        Dim CabJFIF As String
        Dim CabMarcador As String
        Dim CabLongitud As Integer
        Dim iSubBucle As Long
        Dim Unidades As Byte
        Dim entero As Integer
        Dim VersionJFIF As Integer
        Dim PosicionINICab As Long
        Dim TXT As String
        container.Items.Clear()
        'Comprobamos existencia del fichero
        If System.IO.File.Exists(RutaFichero) = False Then
            AnalizarJPGHeader = 99  'El fichero no existe
            Exit Function
        End If
        'Abrimos el fichero
        Try
            FileCad = New IO.FileStream(RutaFichero,
                                                IO.FileMode.Open,
                                                IO.FileAccess.Read,
                                                IO.FileShare.Read)
            FileBin = New IO.BinaryReader(FileCad)
        Catch Manage_err As Exception
            AnalizarJPGHeader = 98  'Error al acceder el fichero
            Exit Function
        End Try
        'Inicializamos las variables
        CabMarcador = ""
        CabJFIF = ""
        'ImageWidth = 0
        'ImageHeight = 0
        'SOI del JPG (Start Of Image)
        If Hex(FileBin.ReadByte) = "FF" Then okProc = True
        If Hex(FileBin.ReadByte) = "D8" Then okProc = True
        If okProc = False Then
            AnalizarJPGHeader = 97  'El fichero no es un JPG
        End If
        Do
            'Localizada una cabecera y calculo de su longitud en bytes
            If Hex(FileBin.ReadByte) = "FF" Then
                CabMarcador = Hex(FileBin.ReadByte)
                PosicionINICab = FileBin.BaseStream.Position - 1
            End If
            CabLongitud = 256 * FileBin.ReadByte
            CabLongitud = CabLongitud + FileBin.ReadByte
            container.Items.Add("Marker:" & CabMarcador & vbTab & "Longitud: " & CabLongitud)
            'Procesamos en función de la cabecera
            Select Case CabMarcador
                Case Is = "E0"  'Marcador Application Marker APP0
                    For iSubBucle = 1 To 4
                        CabJFIF = CabJFIF & FileBin.ReadChar
                    Next iSubBucle
                    If CabJFIF = "JFIF" Then     'Identificador JFIF
                        Letra = FileBin.ReadByte
                        VersionJFIF = 256 * FileBin.ReadByte
                        VersionJFIF = VersionJFIF + FileBin.ReadByte
                        '257:Version 1.1
                        '258:Version 1.2
                        Unidades = FileBin.ReadByte
                        'XDensidad = 256 * FileBin.ReadByte
                        'XDensidad = XDensidad + FileBin.ReadByte
                        'YDensidad = 256 * FileBin.ReadByte
                        'YDensidad = YDensidad + FileBin.ReadByte
                        Application.DoEvents()
                    Else
                        AnalizarJPGHeader = 98
                        Exit Do
                    End If
                Case Is = "EC"  'Marcador APP12
                    Application.DoEvents()
                Case Is = "EE"  'Marcador APP14
                    Application.DoEvents()
                Case Is = "DB"  'Marcador DQT (Define a Quantization Table)
                    Application.DoEvents()
                Case Is = "C0"  'Marcador S0F0 (Baseline DCT)
                    Application.DoEvents()
                    Letra = FileBin.ReadByte
                    'ImageHeight = 256 * FileBin.ReadByte
                    'ImageHeight = ImageHeight + FileBin.ReadByte
                    'ImageWidth = 256 * FileBin.ReadByte
                    'ImageWidth = ImageWidth + FileBin.ReadByte
                    Application.DoEvents()
                    AnalizarJPGHeader = 0
                    Exit Do
                Case Is = "C4"  'Marcador DHT (Define Huffman Table)
                    Application.DoEvents()
                Case Is = "E1"
                    Application.DoEvents()
                    'FileCad.Seek(6, IO.SeekOrigin.Begin)
                    '# Estamos en la posición 6. En este fichero Exif comienza en el 6
                    For iSubBucle = 1 To 4
                        CabJFIF = CabJFIF & FileBin.ReadChar
                    Next iSubBucle
                    If CabJFIF = "Exif" Then
                        'Leemos dos posociones nulos
                        Letra = FileBin.ReadByte
                        Letra = FileBin.ReadByte
                        Dim procesador As String = ""
                        Dim tiffHeaderStart As Integer = FileCad.Position
                        procesador = FileBin.ReadChars(2)
                        'Si es II:Intel. Si es MM: Motorola
                        TXT = Hex(FileBin.ReadByte)
                        TXT = Hex(FileBin.ReadByte) & TXT
                        'Si aquí TXT=002A, la cosa va bien

                        entero = FileBin.ReadInt32
                        'Saltamos a la posición tiffHeaderStart + entero, ya que este offset es relativo al tiffHeaderStart
                        FileCad.Seek(tiffHeaderStart + entero, IO.SeekOrigin.Begin)
                        Dim entryCount As Integer = FileBin.ReadUInt16
                        Dim currentEntry As Integer = 0
                        container.Items.Add("Número de entradas:" & entryCount)
                        For currentEntry = 1 To entryCount
                            container.Items.Add("propiedad:" & FileBin.ReadUInt16 &
                                                ". Posición:" & FileCad.Position - 2 &
                                                ". Valor:")

                            FileCad.Seek(10, IO.SeekOrigin.Current)
                        Next
                        container.Items.Add("--------------------")
                        container.Items.Add("Posición:" & FileCad.Position)
                        'Debug.Print(FileBin.ReadUInt16)
                        Application.DoEvents()
                        'Analizamos una propiedad



                    End If


                    'TXT = ""
                    'For iSubBucle = 4 To CabLongitud
                    ' TXT = TXT & FileBin.ReadChar
                    'Next
                    Application.DoEvents()
                Case Is = "DA"  'Marcador SOS (Start Of Scan)
                    Application.DoEvents()
                Case Else
                    Application.DoEvents()
            End Select
            Application.DoEvents()
            CabMarcador = ""
            FileBin.BaseStream.Position = PosicionINICab + CabLongitud + 1
        Loop
        FileBin.Close()
        FileBin = Nothing
        FileCad.Close()
        FileCad.Dispose()
    End Function

End Class
