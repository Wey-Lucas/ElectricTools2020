'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='
Namespace Engine2

    ''' <summary>
    ''' Permite a abertura de aplicativos, pastas e arquivos
    ''' </summary>
    ''' <remarks></remarks>
    Public Class StartApp

        ''' <summary>
        ''' Abre um arquivo ou diretório
        ''' </summary>
        ''' <param name="FileOrDirectory">Arquivo ou diretorio</param>
        ''' <param name="Modo"></param>
        ''' <remarks></remarks>
        Public Shared Sub Start(ByVal FileOrDirectory As String, ByVal Modo As ProcessWindowStyle)
            Try
                Dim Process As System.Diagnostics.Process = New Process()
                With Process
                    .StartInfo.FileName = FileOrDirectory
                    .StartInfo.WindowStyle = Modo
                    .StartInfo.UseShellExecute = True
                    .Start()
                End With
            Catch ex As System.Exception
                Throw New System.Exception("O recurso '" & FileOrDirectory & "' não foi encontrado")
            End Try
        End Sub

        ''' <summary>
        ''' Abre um programa
        ''' </summary>
        ''' <param name="Application"></param>
        ''' <param name="Arguments"></param>
        ''' <remarks></remarks>
        Public Shared Sub StartApp(Application As String, Optional Arguments As Object = Nothing)
            Process.Start(Application, Arguments)
        End Sub

    End Class

End Namespace


