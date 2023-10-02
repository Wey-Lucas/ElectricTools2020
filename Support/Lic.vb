Imports adoNetExtension.AdoNetConnect
Imports ElectricTools2020.ElectricTools

'=========================================================================================================='
'DESENVOLVIDO POR: LUCAS LEÔNCIO WEY DA SILVA (FORMANDO EM ANALISTA DE SISTEMAS)
'EM:2023
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: wey.lucas1@gmail.com \ (011) 99940-2202
'=========================================================================================================='

Namespace Support

    ''' <summary>
    ''' Classe de segurança para controle de licença
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Lic

        ''' <summary>
        ''' Verifica se a licença existe
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLic() As Boolean
            If MyPlugin.Guid <> "23AD8474-FA4A-47E5-B8A0-471FE95F789C" Then
                MsgBox("Erro de violação." & vbCrLf & "Instale o produto a partir do CD." & vbCrLf & "Se o problema persistir contate o suporte técnico.", MsgBoxStyle.Critical)
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Arquitetura
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eArchitecture
            AutoCAD = 0
            Revit = 1
        End Enum

        ''' <summary>
        ''' Adiciona o registro do comando
        ''' </summary>
        ''' <param name="Architecture">Arquitetura</param>
        ''' <param name="ApplicationName">Nome do aplicativo</param>
        ''' <param name="Command">Comando</param>
        Public Shared Sub Add(Architecture As eArchitecture, ApplicationName As String, Command As String)
            Dim User As String
            User = My.Computer.Name & If(Environment.UserName.Trim <> "", " (" & Environment.UserName & ")", "")
            Using DB As New AdoNet
                With DB
                    .OpenConnection(MyPlugin.StringConnection, AdoNet.Providers.SQL, True)
                    With .Fields
                        .ClearAll()
                        .Add("UserName", User)
                        .Add("Architeture", Architecture.ToString)
                        .Add("Application", ApplicationName.ToUpper.Trim)
                        .Add("Command", Command.ToUpper.Trim)
                        .Add("Date", Date.Now)
                        If .AddRegistry("LiccAccess") = True Then
                            DB.Commit()
                        Else
                            DB.Rollback()
                        End If
                    End With
                    .CloseConnection()
                End With
            End Using
        End Sub

    End Class

End Namespace
