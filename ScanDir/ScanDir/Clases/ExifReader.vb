Imports System.Windows.Forms

Public Class ExifReader
    Inherits ApplicationException
    Implements IDisposable


    Private FileCad As IO.FileStream
    Private FileBin As IO.BinaryReader
    Private JPGOpen As Boolean

    Dim unidades As Integer
    Dim XDensidad As Integer
    Dim YDensidad As Integer
    Dim tiffHeaderStart As Integer
    Dim subIFD0HeaderStart As Integer
    Dim otraHeaderStart As Integer
    Dim procesador As String
    Dim listaPropiedades As New Dictionary(Of ExifTag, Integer)
    Dim listaMetadatos As New Dictionary(Of String, String)

    Public Structure PropiedadEXIF
        Dim Nombre As String
        Dim Valor As String
    End Structure


    Public Sub Dispose() Implements System.IDisposable.Dispose
        If Not FileBin Is Nothing Then
            FileBin.Close()
            FileBin = Nothing
            FileCad.Close()
            FileCad.Dispose()
            FileCad = Nothing
            Debug.Print("Class Disposed")
        End If
    End Sub


    Protected Overrides Sub Finalize()
        If Not FileBin Is Nothing Then
            FileBin.Close()
            FileBin = Nothing
            FileCad.Close()
            FileCad.Dispose()
            FileCad = Nothing
            Debug.Print("Class Finalized")
        End If
        MyBase.Finalize()
    End Sub

    Public Sub New(ByVal Ruta As String)
        If System.IO.File.Exists(Ruta) = False Then
            Throw New ArgumentException("El fichero no existe")
        End If
        Application.DoEvents()
        'Abrimos el fichero
        Try
            FileCad = New IO.FileStream(Ruta, _
                                                IO.FileMode.Open, _
                                                IO.FileAccess.Read, _
                                                IO.FileShare.Read)
            FileBin = New IO.BinaryReader(FileCad)
            JPGOpen = True
            If AnalisisBasico() = False Then
                Throw New ArgumentException("El fichero no tiene formato JPG")
            End If
            AnalizarPropiedades()
        Catch Manage_err As Exception
            Throw Manage_err
        End Try
    End Sub

    Public Sub DamePropiedades()
        If JPGOpen = False Then
            Debug.Print("No hay imagen")
            Exit Sub
        End If
        Debug.Print("Tamaño: " & FileCad.Length)

    End Sub

    Public Function DameMetadatosImagen() As ArrayList
        Dim propiEXIF As New ArrayList()
        Dim pEXIF As PropiedadEXIF
        Dim propData As KeyValuePair(Of String, String)
        For Each propData In listaMetadatos
            pEXIF.Nombre = propData.Key
            pEXIF.Valor = propData.Value
            propiEXIF.Add(pEXIF)
        Next

        Return propiEXIF


    End Function


    Private Function AnalisisBasico() As Boolean

        Dim cabMarker As String
        Dim cabName As String
        Dim okProc As Boolean
        Dim letra As Byte
        Dim VersionJFIF As Integer
        Dim testByte As String
        Dim valorInt As Integer
        Dim PosicionINICab As Integer, CabLongitud As Integer
        Dim typeData As UInt16
        Dim numberOfComponents As UInt32
        Dim arrayBytes As Byte()

        Debug.Print("-------------- Comienzo ------------------")
        cabMarker = ""
        cabName = ""
        'SOI del JPG (Start Of Image)
        If Hex(FileBin.ReadByte) = "FF" Then okProc = True
        If Hex(FileBin.ReadByte) = "D8" Then okProc = True
        If okProc = False Then
            Exit Function
        End If

        'Analizo la cabecera
        Do
            'Localizada una cabecera y calculo de su longitud en bytes
            If Hex(FileBin.ReadByte) = "FF" Then
                cabMarker = Hex(FileBin.ReadByte)
                PosicionINICab = FileBin.BaseStream.Position - 1
            End If
            CabLongitud = 256 * FileBin.ReadByte
            CabLongitud = CabLongitud + FileBin.ReadByte
            Debug.Print("Marker:" & cabMarker & vbTab & "Longitud: " & CabLongitud)
            'Procesamos en función de la cabecera
            Select Case cabMarker
                Case Is = "E0"  'Marcador Application Marker APP0
                    For iSubBucle = 1 To 4
                        cabName = cabName & FileBin.ReadChar
                    Next iSubBucle
                    If cabName = "JFIF" Then     'Identificador JFIF
                        letra = FileBin.ReadByte
                        VersionJFIF = 256 * FileBin.ReadByte
                        VersionJFIF = VersionJFIF + FileBin.ReadByte
                        '257:Version 1.1
                        '258:Version 1.2
                        unidades = FileBin.ReadByte
                        XDensidad = 256 * FileBin.ReadByte
                        XDensidad = XDensidad + FileBin.ReadByte
                        YDensidad = 256 * FileBin.ReadByte
                        YDensidad = YDensidad + FileBin.ReadByte
                        Debug.Print("XDensidad: " & XDensidad)
                        Debug.Print("YDensidad: " & YDensidad)
                    Else
                        Exit Function
                    End If
                Case Is = "EC"  'Marcador APP12
                    Application.DoEvents()
                Case Is = "EE"  'Marcador APP14
                    Application.DoEvents()
                Case Is = "DB"  'Marcador DQT (Define a Quantization Table)
                    Application.DoEvents()
                Case Is = "C0"  'Marcador S0F0 (Baseline DCT)
                    Application.DoEvents()
                    Exit Do
                Case Is = "C4"  'Marcador DHT (Define Huffman Table)
                    Application.DoEvents()
                Case Is = "E1"
                    Application.DoEvents()
                    '# Estamos en la posición 6. En este fichero Exif comienza en el 6
                    For iSubBucle = 1 To 4
                        cabName = cabName & FileBin.ReadChar
                    Next iSubBucle
                    If cabName = "Exif" Then
                        'Leemos dos posiciones nulos
                        letra = FileBin.ReadByte
                        letra = FileBin.ReadByte
                        tiffHeaderStart = FileCad.Position
                        procesador = FileBin.ReadChars(2)
                        'Si es II:Intel. Si es MM: Motorola
                        testByte = Hex(FileBin.ReadByte)
                        testByte = Hex(FileBin.ReadByte) & testByte
                        If testByte <> "02A" Then
                            Exit Function
                        End If
                        valorInt = FileBin.ReadInt32
                        '----------------------------------------------------------
                        CatalogarIFD0(valorInt)
                        If otraHeaderStart > 0 Then
                            '34853=&H8825
                            arrayBytes = GetTagBytes(&H8825, otraHeaderStart, numberOfComponents, typeData)
                            Application.DoEvents()
                            otraHeaderStart = System.BitConverter.ToUInt32(arrayBytes, 0)
                            CatalogarIFD0(otraHeaderStart)
                        End If

                        If subIFD0HeaderStart > 0 Then
                            '34665=&H8769
                            arrayBytes = GetTagBytes(&H8769, subIFD0HeaderStart, numberOfComponents, typeData)
                            subIFD0HeaderStart = System.BitConverter.ToUInt32(arrayBytes, 0)
                            CatalogarIFD0(subIFD0HeaderStart)
                        End If
                        '----------------------------------------------------------
                        Debug.Print("** Elementos almacenados:" & listaPropiedades.Count)
                        Debug.Print("* Posición:" & FileCad.Position)
                        Application.DoEvents()
                    End If
                    Application.DoEvents()
                Case Is = "DA"  'Marcador SOS (Start Of Scan)
                    Application.DoEvents()
                Case Else
                    Application.DoEvents()
            End Select
            Application.DoEvents()
            cabMarker = ""
            FileBin.BaseStream.Position = PosicionINICab + CabLongitud + 1
        Loop
        Debug.Print("-------------- Final ------------------")
        AnalisisBasico = True
    End Function

    Sub CatalogarIFD0(ByVal offsetBegin As Integer)
        Dim numProp As UInt32
        Dim offsetPos As Integer

        'Saltamos a la posición tiffHeaderStart + entero, ya que este offset es relativo al tiffHeaderStart
        FileCad.Seek(tiffHeaderStart + offsetBegin, IO.SeekOrigin.Begin)
        Dim entryCount As Integer = FileBin.ReadUInt16
        Dim currentEntry As Integer = 0
        Debug.Print("** Número de entradas:" & entryCount)
        For currentEntry = 1 To entryCount
            numProp = FileBin.ReadUInt16
            offsetPos = FileCad.Position - 2
            If System.Enum.IsDefined(GetType(ExifTag), numProp) Then
                listaPropiedades.Add(CType(numProp, ExifTag), offsetPos)
                'Debug.Print("* IdPropiedad:" & numProp & _
                '                    ". Posición:" & offsetPos & _
                '                    ". Nombre propiedad:" & [Enum].GetName(GetType(ExifTag), numProp).ToString)
            Else
                Debug.Print("* PROPIEDAD NO REGISTRADA:" & numProp)
                If numProp = 34665 Then
                    subIFD0HeaderStart = offsetPos
                End If
                If numProp = 34853 Then
                    otraHeaderStart = offsetPos
                End If
            End If
            FileCad.Seek(10, IO.SeekOrigin.Current)
        Next

    End Sub




    Private Sub AnalizarPropiedades()

        Dim propData As KeyValuePair(Of ExifTag, Integer)
        Dim arrayBytes As Byte()
        Dim typeData As UInt16
        Dim numberOfComponents As UInt32
        Dim ascii As New System.Text.ASCIIEncoding
        Dim fieldLength As Byte
        Dim cadLog As String
        Dim cadProp As String
        For Each propData In listaPropiedades
            typeData = 0
            numberOfComponents = 0
            arrayBytes = GetTagBytes(propData.Key, propData.Value, numberOfComponents, typeData)
            If propData.Key = ExifTag.GPSLatitude Then
                Application.DoEvents()
            End If
            If numberOfComponents <> 1 And typeData <> 2 And typeData <> 5 Then
                Application.DoEvents()
            End If

            If typeData = 1 Then
                'Unsigned byte
                'Debug.Print(cadLog & ". Valor:" & arrayBytes(0))
                cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
            ElseIf typeData = 2 Then
                'ASCII String
                cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                'listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, ascii.GetString(arrayBytes))
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
            ElseIf typeData = 3 Then
                'Unsigned short 2 bytes
                cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                'Debug.Print(cadLog & ". Valor:" & ToUshort(arrayBytes))
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
            ElseIf typeData = 4 Then
                'Unsigned long 4 bytes
                Debug.Print(cadLog & ". Valor:" & ToUint(arrayBytes))
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, ToUint(arrayBytes))
            ElseIf typeData = 5 Then
                'Unsigned Rational
                'cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                'listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
                If numberOfComponents = 1 Then
                    Debug.Print(cadLog & ". Valor:" & ToURational(arrayBytes))
                    listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, ToURational(arrayBytes))
                Else
                    Dim convertedData() As Double
                    fieldLength = GetTIFFFieldLength(typeData)
                    convertedData = GetArray(arrayBytes, fieldLength, AddressOf ToURational)
                    listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, _
                                       convertedData(0) + convertedData(1) / 60 + convertedData(2) / 3600)
                End If


            ElseIf typeData = 7 Then
                'Undefined. Treat it as an unsigned integer.
                'Debug.Print(cadLog & ". Valor:" & ToUint(arrayBytes))
                cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
                'listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, ToUint(arrayBytes))

            ElseIf typeData = 10 Then
                'Signed Rational
                cadProp = FormatearPropiedades(propData.Key, arrayBytes, typeData, numberOfComponents)
                listaMetadatos.Add([Enum].GetName(GetType(ExifTag), propData.Key).ToString, cadProp)
            Else
                Application.DoEvents()
            End If
        Next
        Application.DoEvents()

    End Sub


    Function FormatearPropiedades(ByVal tagID As Integer, ByVal data As Byte(), _
                                  ByVal Tipodato As UInt16, ByVal numberOfComponents As UInt32) As String

        Dim cadLog As String
        Dim ascii As New System.Text.ASCIIEncoding
        Dim cadArray As String
        Dim valor As UInt16
        Dim valorD As Double

        cadLog = "Extrayendo " & [Enum].GetName(GetType(ExifTag), tagID).ToString & ". Tipodato: " & Tipodato
        Debug.Print(cadLog)
        If Tipodato = 1 Then
            If data.Length = 4 Then
                Return "Versión: " & data(0) & "." & data(1)
            ElseIf Hex(tagID) = &H5 And data(0) = 0 Then
                Return "Nivel del mar"
            Else
                Return data(0)
            End If
        ElseIf Tipodato = 2 Then
            cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
            'cadArray = ascii.GetString(data, 0, 4).ToString.Trim
            If tagID = &H1 Or tagID = &H13 Then
                If data(0) = 78 Then Return "Latitud Norte"
                If data(0) = 83 Then Return "Latitud Sur"
                Return "Reservado"
            End If
            If tagID = &H3 Or tagID = &H15 Then
                If data(0) = 69 Then Return "Longitud Este"
                If data(0) = 87 Then Return "Longitud Oeste"
                Return "Reservado"
            End If
            If tagID = &H9 Then
                If data(0) = 65 Then Return "Measurement in progress"
                If data(0) = 86 Then Return "Measurement Interoperability"
                Return "Reservado"
            End If
            If tagID = &HA Then
                If data(0) = "2" Then Return "2-dimensional measurement"
                If data(0) = "3" Then Return "3-dimensional measurement"
                Return "Reservado"
            End If
            If tagID = &HC Or tagID = &H19 Then
                If data(0) = 75 Then Return "Km/h"
                If data(0) = 77 Then Return "Millas/k"
                If data(0) = 78 Then Return "Nudos"
                Return "Reservado"
            End If
            If tagID = &HE Or tagID = &H10 Or tagID = &H17 Then
                If data(0) = 84 Then Return "True direction"
                If data(0) = 77 Then Return "Magnetic direction"
                Return "Reservado"
            End If
            Return cadArray
        ElseIf Tipodato = 3 Then
            valor = ToUshort(data)
            If tagID = &H8827 Then Return "ISO-" & valor
            If tagID = &HA217 Then 'Sensing method
                If valor = 1 Then Return "Not definied"
                If valor = 2 Then Return "One-chip color area sensor"
                If valor = 3 Then Return "Two-chip color area sensor"
                If valor = 4 Then Return "Three-chip color area sensor"
                If valor = 5 Then Return "Color sequential area sensor"
                If valor = 7 Then Return "Trilinear sensor"
                If valor = 8 Then Return "Color sequential linear sensor"
                Return "Reservado"
            End If
            If tagID = &H8822 Then 'Exposure program
                If valor = 0 Then Return "Not defined"
                If valor = 1 Then Return "Manual"
                If valor = 2 Then Return "Normal program"
                If valor = 3 Then Return "Aperture priority"
                If valor = 4 Then Return "Shutter priority"
                If valor = 5 Then Return "Creative program (biased toward depth of field)"
                If valor = 6 Then Return "Action program (biased toward fast shutter speed)"
                If valor = 7 Then Return "Portrait mode (for closeup photos with the background out of focus)"
                If valor = 8 Then Return "Landscape mode (for landscape photos with the background in focus)"
                Return "Reservado"
            End If
            If tagID = &H9207 Then 'Metering mode
                If valor = 0 Then Return "unknown"
                If valor = 1 Then Return "Average"
                If valor = 2 Then Return "Center Weighted Average"
                If valor = 3 Then Return "Spot"
                If valor = 4 Then Return "MultiSpot"
                If valor = 5 Then Return "Pattern"
                If valor = 6 Then Return "Partial"
                If valor = 255 Then Return "Other"
                Return "Reservado"
            End If
            If tagID = &H9208 Then 'Light source
                If valor = 0 Then Return "unknown"
                If valor = 1 Then Return "Daylight"
                If valor = 2 Then Return "Fluorescent"
                If valor = 3 Then Return "Tungsten (incandescent light)"
                If valor = 4 Then Return "Flash"
                If valor = 9 Then Return "Fine weather"
                If valor = 10 Then Return "Cloudy weather"
                If valor = 11 Then Return "Shade"
                If valor = 12 Then Return "Daylight fluorescent (D 5700 – 7100K)"
                If valor = 13 Then Return "Day white fluorescent (N 4600 – 5400K)"
                If valor = 14 Then Return "Cool white fluorescent (W 3900 – 4500K)"
                If valor = 15 Then Return "White fluorescent (WW 3200 – 3700K)"
                If valor = 17 Then Return "Standard light A"
                If valor = 18 Then Return "Standard light B"
                If valor = 19 Then Return "Standard light C"
                If valor = 20 Then Return "D55"
                If valor = 21 Then Return "D65"
                If valor = 22 Then Return "D75"
                If valor = 23 Then Return "D50"
                If valor = 24 Then Return "ISO studio tungsten"
                If valor = 255 Then Return "ISO studio tungsten"
                Return "Other light source"
            End If
            If tagID = &H9209 Then 'Flash
                If valor = &H0 Then Return "Flash did not fire"
                If valor = &H1 Then Return "Flash fired"
                If valor = &H5 Then Return "Strobe return light not detected"
                If valor = &H7 Then Return "Strobe return light detected"
                If valor = &H9 Then Return "Flash fired, compulsory flash mode"
                If valor = &HD Then Return "Flash fired, compulsory flash mode, return light not detected"
                If valor = &HF Then Return "Flash fired, compulsory flash mode, return light detected"
                If valor = &H10 Then Return "Flash did not fire, compulsory flash mode"
                If valor = &H18 Then Return "Flash did not fire, auto mode"
                If valor = &H19 Then Return "Flash fired, auto mode"
                If valor = &H1D Then Return "Flash fired, auto mode, return light not detected"
                If valor = &H1F Then Return "Flash fired, auto mode, return light detected"
                If valor = &H20 Then Return "No flash function"
                If valor = &H41 Then Return "Flash fired, red-eye reduction mode"
                If valor = &H45 Then Return "Flash fired, red-eye reduction mode, return light not detected"
                If valor = &H47 Then Return "Flash fired, red-eye reduction mode, return light detected"
                If valor = &H49 Then Return "Flash fired, compulsory flash mode, red-eye reduction mode"
                If valor = &H4D Then Return "Flash fired, compulsory flash mode, red-eye reduction mode, return light not detected"
                If valor = &H4F Then Return "Flash fired, compulsory flash mode, red-eye reduction mode, return light detected"
                If valor = &H59 Then Return "Flash fired, auto mode, red-eye reduction mode"
                If valor = &H5D Then Return "Flash fired, auto mode, return light not detected, red-eye reduction mode"
                If valor = &H5F Then Return "Flash fired, auto mode, return light detected, red-eye reduction mode"
                Return "Reservado"
            End If
            If tagID = &H128 Then 'Resolution unit
                If valor = 2 Then Return "Pulgadas"
                If valor = 3 Then Return "Centímetros"
                Return "Sin unidades"
            End If
            If tagID = &HA409 Then 'Saturacion
                If valor = 0 Then Return "Normal"
                If valor = 1 Then Return "Low saturation"
                If valor = 2 Then Return "High saturation"
                Return "Reservado"
            End If
            If tagID = &HA40A Then 'Sharpness
                If valor = 0 Then Return "Normal"
                If valor = 1 Then Return "Soft"
                If valor = 2 Then Return "Hard"
                Return "Reservado"
            End If
            If tagID = &HA408 Then 'Contrast
                If valor = 0 Then Return "Normal"
                If valor = 1 Then Return "Soft"
                If valor = 2 Then Return "Hard"
                Return "Reservado"
            End If
            If tagID = &H103 Then 'Compression
                If valor = 1 Then Return "Uncompressed"
                If valor = 6 Then Return "JPEG compression (thumbnails only)"
                Return "Reservado"
            End If
            If tagID = &H106 Then 'PhotometricInterpretation
                If valor = 2 Then Return "RGB"
                If valor = 6 Then Return "YCbCr"
                Return "Reservado"
            End If
            If tagID = &H112 Then 'Orientation
                If valor = 1 Then Return "The 0th row is at the visual top of the image, and the 0th column is the visual left-hand side."
                If valor = 2 Then Return "The 0th row is at the visual top of the image, and the 0th column is the visual right-hand side."
                If valor = 3 Then Return "The 0th row is at the visual bottom of the image, and the 0th column is the visual right-hand side."
                If valor = 4 Then Return "The 0th row is at the visual bottom of the image, and the 0th column is the visual left-hand side."
                If valor = 5 Then Return "The 0th row is the visual left-hand side of the image, and the 0th column is the visual top."
                If valor = 6 Then Return "The 0th row is the visual right-hand side of the image, and the 0th column is the visual top."
                If valor = 7 Then Return "The 0th row is the visual right-hand side of the image, and the 0th column is the visual bottom."
                If valor = 8 Then Return "The 0th row is the visual left-hand side of the image, and the 0th column is the visual bottom."
                Return "Reservado"
            End If
            If tagID = &H213 Then 'YCbCrPositioning
                If valor = 1 Then Return "centered"
                If valor = 6 Then Return "co-sited"
                Return "Reservado"
            End If
            If tagID = &HA001 Then 'ColorSpace
                If valor = 1 Then Return "sRGB"
                If valor = &HFFFF Then Return "unCalibrated"
                Return "Reservado"
            End If
            If tagID = &HA401 Then 'CustomRendered
                If valor = 0 Then Return "Normal process"
                If valor = 1 Then Return "Custom process"
                Return "Reservado"
            End If
            If tagID = &HA402 Then 'ExposureMode
                If valor = 0 Then Return "Auto exposure"
                If valor = 1 Then Return "Manual exposure"
                If valor = 2 Then Return "Auto bracket"
                Return "Reservado"
            End If
            If tagID = &HA403 Then 'WhiteBalance
                If valor = 0 Then Return "Auto white balance"
                If valor = 1 Then Return "Manual white balance"
                Return "Reservado"
            End If
            If tagID = &HA406 Then 'SceneCaptureType
                If valor = 0 Then Return "Standard"
                If valor = 1 Then Return "Landscape"
                If valor = 2 Then Return "Portrait"
                If valor = 3 Then Return "Night scene"
                Return "Reservado"
            End If
            If tagID = &HA40C Then 'SubjectDistanceRange
                If valor = 0 Then Return "unknown"
                If valor = 1 Then Return "Macro"
                If valor = 2 Then Return "Close view"
                If valor = 3 Then Return "Distant view"
                Return "Reservado"
            End If
            If tagID = &H1E Then 'GPSDifferential
                If valor = 0 Then Return "Measurement without differential correction"
                If valor = 1 Then Return "Differential correction applied"
                Return "Reservado"
            End If
            If tagID = &HA405 Then 'FocalLengthIn35mmFilm
                Return valor & " mm"
            End If
        ElseIf Tipodato = 5 Then
            Dim fieldLength As Byte
            If numberOfComponents = 1 Then
                valorD = ToURational(data)
            Else
                Dim convertedData() As Double
                fieldLength = GetTIFFFieldLength(Tipodato)
                convertedData = GetArray(data, fieldLength, AddressOf ToURational)
                valorD = convertedData(0) + convertedData(1) / 60 + convertedData(2) / 3600
            End If
            If tagID = &H9202 Then Return "F/" & Math.Round(Math.Pow(Math.Sqrt(2), valorD), 2).ToString 'ApertureValue
            If tagID = &H9205 Then Return "F/" & Math.Round(Math.Pow(Math.Sqrt(2), valorD), 2).ToString 'MaxApertureValue
            If tagID = &H920A Then Return valorD.ToString & " mm" 'FocalLength
            If tagID = &H829D Then Return "F/" & valorD.ToString 'F-number

            If tagID = &H11A Then Return valorD.ToString 'Xresolution
            If tagID = &H11B Then Return valorD.ToString 'Yresolution
            If tagID = &H829A Then Return valorD.ToString & " sec." 'ExposureTime
            If tagID = &H2 Then Return Math.Round(valorD, 5).ToString & " grados" 'GPSLatitude
            If tagID = &H4 Then Return Math.Round(valorD, 5).ToString & " grados" 'GPSLongitude
            If tagID = &H6 Then Return valorD.ToString & " m." 'GPSAltitude
            If tagID = &HA404 Then Return valorD.ToString 'Digital Zoom Ratio
            If tagID = &HB Then Return valorD.ToString 'GPSDOP
            If tagID = &HD Then Return valorD.ToString 'GPSSpeed
            If tagID = &HF Then Return valorD.ToString 'GPSTrack

            If tagID = &H11 Then Return valorD.ToString 'GPSImgDir
            If tagID = &H14 Then Return valorD.ToString 'GPSDestLatitude
            If tagID = &H16 Then Return valorD.ToString 'GPSDestLongitude
            If tagID = &H18 Then Return valorD.ToString 'GPSBearing
            If tagID = &H1A Then Return valorD.ToString 'GPSDestDistance
            If tagID = &H7 Then Return valorD.ToString 'GPSTimestamp
        ElseIf Tipodato = 7 Then
            If tagID = &HA300 Then 'FileSource
                If data(0) = 3 Then : Return "DSC" : Else : Return "Reserved" : End If
            End If
            If tagID = &HA301 Then 'SceneType
                If data(0) = 1 Then : Return "A directly photographed image" : Else : Return "Reserved" : End If
            End If
            If tagID = &H9000 Then 'Exif Version
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
            End If
            If tagID = &HA000 Then 'Flashpix Version
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
                Return cadArray
            End If
            If tagID = &H9101 Then 'ComponentsConfiguration
                Application.DoEvents()
            End If
            If tagID = &H927C Then 'MakerNote
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
                Return cadArray
            End If
            If tagID = &H9286 Then 'UserComment
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
                Return cadArray
            End If
            If tagID = &H1B Then 'GPS Processing Method
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
            End If
            If tagID = &H1C Then 'GPS Area Info
                Application.DoEvents()
                cadArray = ascii.GetString(data, 0, numberOfComponents).ToString.Trim
            End If
        ElseIf Tipodato = 10 Then
            valorD = ToRational(data)
            If tagID = &H9201 Then 'ShutterSpeedValue
                Return "1/" & Math.Round(Math.Pow(2, valorD), 2).ToString()
            End If
            If tagID = &H9203 Then 'BrightnessValue
                Return Math.Round(valorD, 4).ToString()
            End If
            If tagID = &H9204 Then 'ExposureBiasValue
                Return Math.Round(valorD, 2).ToString() + " eV"
            End If




        End If









    End Function



    Function GetTagBytes(ByVal tagID As Integer, ByVal tagOffset As Integer, _
                            ByRef numberOfComponents As UInt32, ByRef tiffDataType As UInt16) As Byte()

        'Nos colocamos en posición y comprobamos que aquí se encuentra el tag que buscamos
        Dim offsetAddress As UInt16
        Dim dataSize As Integer

        FileCad.Seek(tagOffset, IO.SeekOrigin.Begin)
        If FileBin.ReadUInt16 <> tagID Then
            Throw New Exception("Tag no encontrado en posición offset")
        End If
        Application.DoEvents()
        tiffDataType = FileBin.ReadUInt16
        numberOfComponents = FileBin.ReadUInt32
        GetTagBytes = FileBin.ReadBytes(4)
        dataSize = numberOfComponents * GetTIFFFieldLength(tiffDataType)
        If dataSize > 4 Then
            Application.DoEvents()
            offsetAddress = System.BitConverter.ToUInt16(GetTagBytes, 0)
            Dim ArrayBytes() As Byte
            FileCad.Seek(offsetAddress + tiffHeaderStart, IO.SeekOrigin.Begin)
            'ArrayBytes = FileBin.ReadBytes(dataSize)
            'Return ArrayBytes
            Return FileBin.ReadBytes(dataSize)
        Else
            Return GetTagBytes
            Application.DoEvents()
        End If


    End Function


    Function GetTIFFFieldLength(ByVal valor As UInt16) As Byte

        Select Case valor
            Case 1
                Return 1
            Case 2
                Return 1
            Case 6
                Return 1
            Case 3
                Return 2
            Case 8
                Return 2
            Case 4
                Return 4
            Case 7
                Return 4
            Case 9
                Return 4
            Case 11
                Return 4
            Case 5
                Return 8
            Case 10
                Return 8
            Case 12
                Return 8
            Case Else
                Throw New Exception(String.Format("Unknown TIFF datatype: {0} " & valor))
        End Select


    End Function




    ''' <summary>
    ''' Convert 8 bytes to an unsigned rational using the current byte aligns.
    ''' </summary>
    ''' <param name="arrayBytes"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ToURational(ByVal arrayBytes() As Byte) As Double

        Dim numerador(3) As Byte
        Dim denominador(3) As Byte
        Array.Copy(arrayBytes, numerador, 4)
        Array.Copy(arrayBytes, 4, denominador, 0, 4)
        Return System.BitConverter.ToUInt32(numerador, 0) / System.BitConverter.ToUInt32(denominador, 0)

    End Function

    ''' <summary>
    '''  Converts 8 bytes to a signed rational using the current byte aligns.
    '''  A TIFF rational contains 2 4-byte integers, the first of which is the numerator, and the second 
    '''  of which is the denominator.
    ''' </summary>
    ''' <param name="arrayBytes"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ToRational(ByVal arrayBytes() As Byte) As Double

        Dim numerador(3) As Byte
        Dim denominador(3) As Byte
        Array.Copy(arrayBytes, numerador, 4)
        Array.Copy(arrayBytes, 4, denominador, 0, 4)
        Return System.BitConverter.ToInt32(numerador, 0) / System.BitConverter.ToInt32(denominador, 0)

    End Function


    Function ToUint(ByVal arrayBytes() As Byte) As UInt32
        Return System.BitConverter.ToUInt32(arrayBytes, 0)
    End Function

    Function ToUshort(ByVal arrayBytes() As Byte) As UInt32
        Return System.BitConverter.ToUInt16(arrayBytes, 0)
    End Function

    Private Delegate Function ConverterMethod(ByVal data() As Byte)

    'private Array GetArray<T>(byte[] data, int elementLengthBytes, ConverterMethod<T> reader)
    Private Function GetArray(ByVal data() As Byte, ByVal elementLengthBytes As Integer, ByVal reader As ConverterMethod) As Array
        Dim convertedData As Array
        convertedData = Array.CreateInstance(GetType(Double), CType(data.Length / elementLengthBytes, Integer))
        Dim buffer(elementLengthBytes) As Byte
        For iElem As Integer = 0 To (data.Length / elementLengthBytes) - 1
            Array.Copy(data, iElem * elementLengthBytes, buffer, 0, elementLengthBytes)
            'convertedData.SetValue(ToURational(buffer), iElem)
            convertedData.SetValue(reader(buffer), iElem)
        Next
        Return convertedData
    End Function


End Class
