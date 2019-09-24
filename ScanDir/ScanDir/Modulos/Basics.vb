Module Basics
    Dim FicheroLogger As String = My.Application.Info.DirectoryPath & "\logger.log"
    Public Const AplicacionTitulo As String = "Scandir 2007"



    Function GenerarLOG(ByVal Frase As String) As Boolean


        Dim sw As New System.IO.StreamWriter(FicheroLogger, True)

        Dim cadFechaInsert As String = Now.Year & "-" & _
                                        String.Format("{0:00}", CInt(Now.Month.ToString)) & "-" & _
                                        String.Format("{0:00}", CInt(Now.Day.ToString)) & " " & _
                                        String.Format("{0:00}", CInt(Now.Hour.ToString)) & ":" & _
                                        String.Format("{0:00}", CInt(Now.Minute.ToString))

        sw.WriteLine(cadFechaInsert & " # " & Frase)
        sw.Close()
        sw.Dispose()
        sw = Nothing

    End Function
End Module
