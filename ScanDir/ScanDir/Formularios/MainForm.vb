Imports System.Windows.Forms

Public Class MainForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        FolderBrowserDialog1.ShowDialog()
        TextBox1.Text = FolderBrowserDialog1.SelectedPath

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim ProcEscaneo As ScanDir
        Dim ListaExtensiones() As String
        If TextBox1.Text.Trim = "" Then
            ErrorProvider1.SetError(TextBox1, "Debe seleccionar un directorio")
            Exit Sub
        Else
            ErrorProvider1.SetError(TextBox1, "")
        End If
        If CheckBox1.Checked = True Then
            ReDim ListaExtensiones(0 To 2)
            ListaExtensiones(0) = ".jpg"
            ListaExtensiones(1) = ".gif"
            ListaExtensiones(2) = ".tif"
            'Ordenamos el array para las busquedfas
            Array.Sort(ListaExtensiones)
        Else

            'ReDim ListaExtensiones(0 To 1)
            'ListaExtensiones(0) = ".ecw"
            'ListaExtensiones(1) = ".ECW"
            ListaExtensiones = Nothing
        End If
        '----------------------------------------------------
        ToolStripStatusLabel1.Text = "Procesando..."
        ProcEscaneo = New ScanDir
        If CheckBox2.Checked = True Then
            ProcEscaneo.ProcesarJPGcabeceras = True
        Else
            ProcEscaneo.ProcesarJPGcabeceras = False
        End If



        Me.Cursor = Cursors.WaitCursor
        If ProcEscaneo.ListarFicherosDIR(TextBox1.Text, ListaExtensiones) = 0 Then
            ToolStripStatusLabel1.Text = "Terminado. " &
                                "Directorios: " & ProcEscaneo.NumeroDirectorios & ". " &
                                "Ficheros: " & ProcEscaneo.NumeroFicheros
            Me.Cursor = Cursors.Default
            If MessageBox.Show("Proceso terminado." & vbCrLf & "Generada: " & ProcEscaneo.DIRBase & vbCrLf & "¿Desea abrir la BD?",
                            "SCANDIR", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                Process.Start(My.Application.Info.DirectoryPath.ToString)
            End If

        Else
            Me.Cursor = Cursors.Default
            ToolStripStatusLabel1.Text = "Terminado. Se han producido errores"
        End If
        ProcEscaneo = Nothing

    End Sub

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        ToolStripStatusLabel1.Text = "Preparado"
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim lector As ExifReader
        Dim listaPropiedades As ArrayList
        Dim elementoLV As ListViewItem

        ListView1.Items.Clear()
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Propiedad", 230, HorizontalAlignment.Left)
        ListView1.Columns.Add("Valor", 230, HorizontalAlignment.Right)
        ListView1.GridLines = True
        ListView1.FullRowSelect = True
        ListView1.View = View.Details


        Try
            lector = New ExifReader("C:\Documents and Settings\e2molin\" &
                                            "Mis documentos\Mis imágenes\Fotos\Avila2008\DSC04946.JPG")
            listaPropiedades = lector.DameMetadatosImagen()
            For Each propiedad As ExifReader.PropiedadEXIF In listaPropiedades
                Debug.Print("Muestro: " & propiedad.Nombre)
                If Not propiedad.Valor = Nothing Then
                    elementoLV = New ListViewItem
                    elementoLV.Text = propiedad.Nombre
                    elementoLV.SubItems.Add(propiedad.Valor.ToString)
                    ListView1.Items.Add(elementoLV)
                    elementoLV = Nothing
                End If
            Next
            lector.Dispose()
        Catch ex As Exception
            Debug.Print(ex.Message)

        End Try


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim valor As Integer = 0
        Dim respuesta As String = ""

        If ConectarBD(TiposBase.SQLlite, "", 0, "", "", "", My.Application.Info.DirectoryPath & "\dirdisk.db3") = True Then
            ObtenerEscalarNumerico("select count(*) from rutasfiles", valor)
            ObtenerEscalarTexto("select count(*) from rutasfiles", respuesta)
            DesconectarBD()
        End If



        Application.DoEvents()
    End Sub
End Class
