'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace Engine2

    ''' <summary>
    ''' Invoca comandos do AutoCad e permite obter as entidades criadas
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class CommandSlave

        'Variáveis
        Private WithEvents _AcadApplication As Autodesk.AutoCAD.Interop.AcadApplication
        Private WithEvents _AcadDocument As Document
        Private WithEvents _AcadDatabase As Database
        Private _DBObjectCollection As DBObjectCollection
        Private _AcadEntityType As ArrayList
        Private _CommandInvoke As String
        Private _IsActiveCommand As Boolean
        Private _ReCommandAction As eReCommandAction
        Private _Tag As Object

        ''' <summary>
        ''' Ações para reentrada de comandos
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eReCommandAction
            EndLastCommandAndCreateEntitys = 0
            AbortLastCommand = 1
            AbortNewCommand = 3
        End Enum

        ''' <summary>
        ''' Comando a ser iniciado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="GlobalCommandName"></param>
        ''' <remarks></remarks>
        Public Event CommandInvoke(sender As Object, GlobalCommandName As String, ByRef e As InvokeEventArgs)

        ''' <summary>
        ''' Comando iniciado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="GlobalCommandName"></param>
        ''' <remarks></remarks>
        Public Event CommandStarted(sender As Object, GlobalCommandName As String)

        ''' <summary>
        ''' Comando finalizado com sucesso (Criou entidades)
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="GlobalCommandName"></param>
        ''' <param name="DBObjectCollection"></param>
        ''' <remarks></remarks>
        Public Event CommandCreateEntity(sender As Object, GlobalCommandName As String, DBObjectCollection As DBObjectCollection)

        ''' <summary>
        ''' Comando cancelado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="GlobalCommandName"></param>
        ''' <remarks></remarks>
        Public Event CommandCancelled(sender As Object, GlobalCommandName As String)

        ''' <summary>
        ''' Comando não encontrado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="GlobalCommandName"></param>
        ''' <remarks></remarks>
        Public Event UnknownCommand(sender As Object, GlobalCommandName As String)

        ''' <summary>
        ''' Configura o modo como o sistema irá tratar a reentrada de comandos
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ReCommandAction As eReCommandAction
            Get
                Return Me._ReCommandAction
            End Get
            Set(value As eReCommandAction)
                Me._ReCommandAction = value
            End Set
        End Property

        ''' <summary>
        ''' Retorna se existe comando ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsActiveCommand As Boolean
            Get
                Return Me._IsActiveCommand
            End Get
        End Property

        ''' <summary>
        ''' Retorna se existe comando ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ActiveGlobalCommandName As String
            Get
                Return If(Me._CommandInvoke <> "*", Me._CommandInvoke, "")
            End Get
        End Property

        ''' <summary>
        ''' Para uso do usuário
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Tag As Object
            Get
                Return Me._Tag
            End Get
            Set(value As Object)
                Me._Tag = value
            End Set
        End Property

        ''' <summary>
        ''' Adiciona itens na biblioteca
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AcadDatabase_ObjectAppended(sender As Object, e As ObjectEventArgs)
            If Me._DBObjectCollection.Contains(e.DBObject) = False And Me._AcadEntityType.Contains(e.DBObject.GetType.Name.ToUpper.Trim) = True Then
                Me._DBObjectCollection.Add(e.DBObject)
            End If
        End Sub

        ''' <summary>
        ''' Remove itens da biblioteca
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AcadDatabase_ObjectErased(sender As Object, e As ObjectErasedEventArgs)
            If Me._DBObjectCollection.Contains(e.DBObject) = True Then
                Me._DBObjectCollection.Remove(e.DBObject)
            End If
        End Sub

        ''' <summary>
        ''' Inicio do comando
        ''' </summary>
        ''' <param name="CommandName"></param>
        ''' <remarks></remarks>
        Private Sub _AcadApplication_BeginCommand(CommandName As String)
            If CommandName = Me._CommandInvoke Then
                Me._IsActiveCommand = True
                RaiseEvent CommandStarted(Me, CommandName)
            End If
        End Sub

        ''' <summary>
        ''' Fim do comando
        ''' </summary>
        ''' <param name="CommandName"></param>
        ''' <remarks></remarks>
        Private Sub _AcadApplication_EndCommand(CommandName As String)
            If CommandName = Me._CommandInvoke Then
                Me._AcadEntityType = Nothing
                Me._CommandInvoke = "*"
                RemoveHandler Me._AcadApplication.BeginCommand, AddressOf Me._AcadApplication_BeginCommand
                RemoveHandler Me._AcadApplication.EndCommand, AddressOf Me._AcadApplication_EndCommand
                RemoveHandler Me._AcadDocument.CommandCancelled, AddressOf Me._AcadDocument_CommandCancelled
                RemoveHandler Me._AcadDocument.UnknownCommand, AddressOf Me._AcadDocument_UnknownCommand
                RemoveHandler Me._AcadDatabase.ObjectAppended, AddressOf Me._AcadDatabase_ObjectAppended
                RemoveHandler Me._AcadDatabase.ObjectErased, AddressOf Me._AcadDatabase_ObjectErased
                Me._IsActiveCommand = False
                If Me._DBObjectCollection.Count > 0 Then
                    RaiseEvent CommandCreateEntity(Me, CommandName, Me._DBObjectCollection)
                Else
                    RaiseEvent CommandCancelled(Me, CommandName)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Comando cancelado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AcadDocument_CommandCancelled(sender As Object, e As CommandEventArgs) 'Handles _AcadDocument.CommandCancelled
            Me._AcadApplication_EndCommand(e.GlobalCommandName)
        End Sub

        ''' <summary>
        ''' Comando não encontrado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AcadDocument_UnknownCommand(sender As Object, e As UnknownCommandEventArgs) 'Handles _AcadDocument.UnknownCommand
            If e.GlobalCommandName = Me._CommandInvoke Then
                Me._AcadDocument.SendStringToExecute("(command)(princ)" & vbLf, True, False, False)
                Me._AcadDocument.Editor.WriteMessage(vbLf & "O comando '" & e.GlobalCommandName & "' não respondeu à solicitação no tempo previsto, repita a operação" & vbLf)
                Me._AcadEntityType = Nothing
                Me._CommandInvoke = "*"
                RemoveHandler Me._AcadApplication.BeginCommand, AddressOf Me._AcadApplication_BeginCommand
                RemoveHandler Me._AcadApplication.EndCommand, AddressOf Me._AcadApplication_EndCommand
                RemoveHandler Me._AcadDocument.CommandCancelled, AddressOf Me._AcadDocument_CommandCancelled
                RemoveHandler Me._AcadDocument.UnknownCommand, AddressOf Me._AcadDocument_UnknownCommand
                RemoveHandler Me._AcadDatabase.ObjectAppended, AddressOf Me._AcadDatabase_ObjectAppended
                RemoveHandler Me._AcadDatabase.ObjectErased, AddressOf Me._AcadDatabase_ObjectErased
                Me._DBObjectCollection.Clear()
                Me._IsActiveCommand = False
                RaiseEvent UnknownCommand(Me, e.GlobalCommandName)
            End If
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="GlobalCommandName">Comando</param>
        ''' <param name="AcadEntityType">Nome da(s) entidade(s) a serem geradas (GetType.Name)</param>
        ''' <param name="Activate">Activate</param>
        ''' <param name="WrapUpInactiveDoc">WrapUpInactiveDoc</param>
        ''' <param name="EchoCommand">EchoCommand</param>
        ''' <param name="Tag">Para uso do usuário</param>
        ''' <remarks></remarks>
        Public Sub Invoke(GlobalCommandName As String, AcadEntityType As ArrayList, Optional Activate As Boolean = True, Optional WrapUpInactiveDoc As Boolean = False, Optional EchoCommand As Boolean = False, Optional Tag As Object = Nothing)
            Try
                Me._Tag = Tag
                If Me._IsActiveCommand = True Then
                    Select Case Me._ReCommandAction
                        Case eReCommandAction.AbortNewCommand
                            Me._AcadDocument.Editor.WriteMessage(vbLf & "O comando '" & Me._CommandInvoke & "' está em uso, conclua o comando e repita a operação." & vbLf)
                            Exit Sub
                        Case eReCommandAction.EndLastCommandAndCreateEntitys
                            Me._AcadApplication_EndCommand(Me._CommandInvoke)
                        Case eReCommandAction.AbortLastCommand
                            While Me._DBObjectCollection.Count <> 0
                                Engine2.Edit.EraseEntity(Engine2.ConvertObject.DBObjectToEntity(Me._DBObjectCollection.Item(0)))
                            End While
                            Me._DBObjectCollection.Clear()
                            Me._AcadDocument.Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                            Me._AcadApplication_EndCommand(Me._CommandInvoke)
                    End Select
                End If
                Dim ComParts() As String = GlobalCommandName.Split(" ")
                Me._CommandInvoke = ComParts(0).Replace(".", "").ToString.Replace("_", "").ToString
                Dim InvokeEventArgs As New InvokeEventArgs(GlobalCommandName)
                RaiseEvent CommandInvoke(Me, GlobalCommandName, InvokeEventArgs)
                If InvokeEventArgs.Cancel = False Then
                    Me._AcadEntityType = New ArrayList
                    For Each Type As String In AcadEntityType
                        If Me._AcadEntityType.Contains(Type.ToUpper.Trim) = False Then
                            Me._AcadEntityType.Add(Type.ToUpper.Trim)
                        End If
                    Next
                    Me._AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
                    Me._AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Me._AcadDatabase = Me._AcadDocument.Database
                    AddHandler Me._AcadApplication.BeginCommand, AddressOf Me._AcadApplication_BeginCommand
                    AddHandler Me._AcadApplication.EndCommand, AddressOf Me._AcadApplication_EndCommand
                    AddHandler Me._AcadDocument.CommandCancelled, AddressOf Me._AcadDocument_CommandCancelled
                    AddHandler Me._AcadDocument.UnknownCommand, AddressOf Me._AcadDocument_UnknownCommand
                    AddHandler Me._AcadDatabase.ObjectAppended, AddressOf Me._AcadDatabase_ObjectAppended
                    AddHandler Me._AcadDatabase.ObjectErased, AddressOf Me._AcadDatabase_ObjectErased
                    Me._DBObjectCollection = New DBObjectCollection
                    With Me._AcadDocument
                        .SendStringToExecute("(command)(princ)" & vbLf, True, False, False)
                        .Window.Focus()
                        If InvokeEventArgs.Message <> "" Then
                            .Editor.WriteMessage(vbLf & InvokeEventArgs.Message & " ")
                        End If
                        .SendStringToExecute(GlobalCommandName.ToUpper.Trim & vbLf, Activate, WrapUpInactiveDoc, EchoCommand)
                    End With
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Command.Invoke, motivo: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ReComandAction">Configura o modo como o sistema irá tratar a reentrada de comandos</param>
        ''' <remarks></remarks>
        Public Sub New(Optional ReComandAction As eReCommandAction = eReCommandAction.AbortLastCommand)
            Me._Tag = Nothing
            Me._ReCommandAction = ReComandAction
        End Sub

    End Class

    ''' <summary>
    ''' Argumentos do evento Invoke
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class InvokeEventArgs

        'Variável
        Private _Cancel As Boolean
        Private _GlobalCommandName As String
        Private _Message As String

        ''' <summary>
        ''' Comando solicitado
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property GlobalCommandName As String
            Get
                Return Me._GlobalCommandName
            End Get
        End Property


        ''' <summary>
        ''' Determina se o evento deve cancelar o processo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Cancel As Boolean
            Get
                Return Me._Cancel
            End Get
            Set(value As Boolean)
                Me._Cancel = value
            End Set
        End Property

        ''' <summary>
        ''' Mensagem
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Message As String
            Get
                Return Me._Message
            End Get
            Set(value As String)
                Me._Message = value
            End Set
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(GlobalCommandName As String)
            Me._Cancel = False
            Me._GlobalCommandName = GlobalCommandName
            Me._Message = ""
        End Sub

    End Class

End Namespace


