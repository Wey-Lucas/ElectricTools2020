Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Text
Imports System.Linq
Imports System.Runtime.ExceptionServices

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2016
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Armazena e gerencia a coleção de reatores
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Reactors : Inherits List(Of Reactor) : Implements IDisposable

#Region "IDisposable  "

        ''' <summary>
        ''' Habilita a destruição
        ''' </summary>
        ''' <remarks></remarks>
        Private DisposedValue As Boolean

        ''' <summary>
        ''' Destroi itens da classe
        ''' </summary>
        ''' <param name="Disposing"></param>
        ''' <remarks></remarks>
        Protected Sub Dispose(Disposing As Boolean)
            If Not Me.DisposedValue Then
                If Disposing Then
                    Me._GlobalCommandName = Nothing
                    If IsNothing(Me._ActiveCommandWillStart) = False Then
                        RemoveHandler Me._Document.CommandWillStart, AddressOf Me.Reactors_Document_CommandWillStart
                    End If
                    If IsNothing(Me._ActiveCommandCancelled) = False Then
                        RemoveHandler Me._Document.CommandCancelled, AddressOf Me.Reactors_Document_CommandCancelled
                    End If
                    If IsNothing(Me._ActiveCommandEnded) = False Then
                        RemoveHandler Me._Document.CommandEnded, AddressOf Me.Reactors_Document_CommandEnded
                    End If
                    If IsNothing(Me._ChangedProcessEventArgs) = False Then
                        Me._ChangedProcessEventArgs.Dispose()
                        Me._ChangedProcessEventArgs = Nothing
                    End If
                    If IsNothing(Me._ClonedProcessEventArgs) = False Then
                        Me._ClonedProcessEventArgs.Dispose()
                        Me._ClonedProcessEventArgs = Nothing
                    End If
                    If IsNothing(Me._ErasedProcessEventArgs) = False Then
                        Me._ErasedProcessEventArgs.Dispose()
                        Me._ErasedProcessEventArgs = Nothing
                    End If
                    If IsNothing(Me._ExplodedProcessEventArgs) = False Then
                        Me._ExplodedProcessEventArgs.Dispose()
                        Me._ExplodedProcessEventArgs = Nothing
                    End If
                    If IsNothing(Me._RestrictProcessEventArgs) = False Then
                        Me._RestrictProcessEventArgs.Dispose()
                        Me._RestrictProcessEventArgs = Nothing
                    End If
                    Me._ChangeCommands.Clear()
                    Me._CloneCommands.Clear()
                    Me._ExplodeCommands.Clear()
                    Me._EraseCommands.Clear()
                    Me._RestrictCommands.Clear()
                    Me._UserMonitoredCommands.Clear()
                    Me._SystemMonitoredCommands.Clear()
                    Me._VetoCommand.Dispose()
                    MyBase.Clear()
                    MyBase.Finalize()
                End If
            End If
            Me.DisposedValue = True
        End Sub

        ''' <summary>
        ''' Destruidor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

        ''' <summary>
        ''' Detecta mudanças a nível de documento
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents _DocumentCollection As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager

        ''' <summary>
        ''' Evento que indica o fim de uma reação
        ''' </summary>
        ''' <param name="ReactorEvents">O tipo de reação finalizada</param>
        ''' <remarks></remarks>
        Public Event EndReaction(ReactorEvents As Reactor.eReactorEvents)

        ''' <summary>
        ''' Evento de erro do gerenciador de reação
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="Message">Mensagem</param>
        ''' <remarks></remarks>
        Public Event ReactorsError(ByRef Reactors As Reactors, Message As String)

        ''' <summary>
        ''' Evento que informa quando um comando qualquer é chamado (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event CommandStarting(ByRef Reactors As Reactors, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento que informa quando um comando monitorado pelo usuário é disparado
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="GlobalCommandname">GlobalCommandname</param>
        ''' <remarks></remarks>
        Public Event MonitoredCommandStarted(ByRef Reactors As Reactors, GlobalCommandName As String)

        ''' <summary>
        ''' Evento que informa quando um comando monitorado pelo usuário é finalizado
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="GlobalCommandname">GlobalCommandname</param>
        ''' <remarks></remarks>
        Public Event MonitoredCommandEnded(ByRef Reactors As Reactors, GlobalCommandName As String)

        ''' <summary>
        ''' Evento que dispara quando um comando monitorado pelo usuário é cancelado
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="GlobalCommandname">GlobalCommandname</param>
        ''' <remarks></remarks>
        Public Event MonitoredCommandCanceled(ByRef Reactors As Reactors, GlobalCommandName As String)

        ''' <summary>
        ''' Evento que informa quando um reator é adicionado
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="Reactor">Reator</param>
        ''' <remarks></remarks>
        Public Event ReactorAdded(ByRef Reactors As Reactors, Reactor As Reactor)

        ''' <summary>
        ''' Evento que informa quando um reator é removido
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="Reactor">Reator</param>
        ''' <remarks></remarks>
        Public Event ReactorRemoved(ByRef Reactors As Reactors, Reactor As Reactor)

        ''' <summary>
        ''' Evento de pré inicialização da reação de clonagem (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ObjectIds">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartingClonedReactions(ByRef Reactors As Reactors, ByRef ObjectIds As List(Of ObjectId), ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de inicialização da reação de clonagem
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ClonedProcessEventArgs">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartedClonedReactions(ByRef Reactors As Reactors, ByRef ClonedProcessEventArgs As ClonedProcessEventArgs, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de clonagem
        ''' </summary>
        ''' <param name="Reactors">Reactos</param>
        ''' <param name="ClonedProcessEventArgs">ClonedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event ClonedReactions(ByRef Reactors As Reactors, ByRef ClonedProcessEventArgs As ClonedProcessEventArgs)

        ''' <summary>
        ''' Evento de finalização da reação de clonagem
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ClonedProcessEventArgs">ClonedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event EndedClonedReactions(ByRef Reactors As Reactors, ByRef ClonedProcessEventArgs As ClonedProcessEventArgs)

        ''' <summary>
        ''' Evento de pré inicialização da reação de exclusão (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ObjectIds">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartingErasedReactions(ByRef Reactors As Reactors, ByRef ObjectIds As List(Of ObjectId), ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de inicialização da reação de exclusão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ErasedProcessEventArgs">ErasedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartedErasedReactions(ByRef Reactors As Reactors, ByRef ErasedProcessEventArgs As ErasedProcessEventArgs, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de exclusão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ErasedProcessEventArgs">ErasedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event ErasedReactions(ByRef Reactors As Reactors, ByRef ErasedProcessEventArgs As ErasedProcessEventArgs)

        ''' <summary>
        ''' Evento de finalização da reação de exclusão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ErasedProcessEventArgs">ErasedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event EndedErasedReactions(ByRef Reactors As Reactors, ByRef ErasedProcessEventArgs As ErasedProcessEventArgs)

        ''' <summary>
        ''' Evento de pré inicialização da reação de explosão (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ObjectIds">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartingExplodedReactions(ByRef Reactors As Reactors, ByRef ObjectIds As List(Of ObjectId), ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de inicialização da reação de explosão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ExplodedProcessEventArgs">ExplodedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartedExplodedReactions(ByRef Reactors As Reactors, ByRef ExplodedProcessEventArgs As ExplodedProcessEventArgs, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de explosão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ExplodedProcessEventArgs">ExplodedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event ExplodedReactions(ByRef Reactors As Reactors, ByRef ExplodedProcessEventArgs As ExplodedProcessEventArgs)

        ''' <summary>
        ''' Evento de finalização da reação de explosão
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ExplodedProcessEventArgs">ExplodedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event EndedExplodedReactions(ByRef Reactors As Reactors, ByRef ExplodedProcessEventArgs As ExplodedProcessEventArgs)

        ''' <summary>
        ''' Evento de pré inicialização da reação de modificação (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ObjectIds">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartingChangedReactions(ByRef Reactors As Reactors, ByRef ObjectIds As List(Of ObjectId), ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de inicialização da reação de modificação
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ChangedProcessEventArgs">ChangedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartedChangedReactions(ByRef Reactors As Reactors, ByRef ChangedProcessEventArgs As ChangedProcessEventArgs, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de modificação
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ChangedProcessEventArgs">ChangedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event ChangedReactions(ByRef Reactors As Reactors, ByRef ChangedProcessEventArgs As ChangedProcessEventArgs)

        ''' <summary>
        ''' Evento de finalização da reação de modificação
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ChangedProcessEventArgs">ChangedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event EndedChangedReactions(ByRef Reactors As Reactors, ByRef ChangedProcessEventArgs As ChangedProcessEventArgs)

        ''' <summary>
        ''' Evento de pré inicialização da reação de seleção de um item monitorado por comando restrito (Antes do comando iniciar)
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="ObjectIds">ClonedProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartingRestrictReactions(ByRef Reactors As Reactors, ByRef ObjectIds As List(Of ObjectId), ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento de inicialização da reação de seleção de um item monitorado por comando restrito
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="RestrictProcessEventArgs">RestrictProcessEventArgs</param>
        ''' <param name="Cancel">Determina se o processo será concelado</param>
        ''' <remarks></remarks>
        Public Event StartedRestrictReactions(ByRef Reactors As Reactors, ByRef RestrictProcessEventArgs As RestrictProcessEventArgs, ByRef Cancel As Boolean)

        ''' <summary>
        ''' Evento que detecta a seleção de um item monitorado por comando restrito
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="RestrictProcessEventArgs">RestrictProcessEventArgsas</param>
        ''' <remarks></remarks>
        Public Event RestrictReactions(ByRef Reactors As Reactors, ByRef RestrictProcessEventArgs As RestrictProcessEventArgs)

        ''' <summary>
        ''' Evento de finalização da reação de seleção de um item monitorado por comando restrito
        ''' </summary>
        ''' <param name="Reactors">Reactors</param>
        ''' <param name="RestrictProcessEventArgs">RestrictProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Event EndedRestrictReactions(ByRef Reactors As Reactors, ByRef RestrictProcessEventArgs As RestrictProcessEventArgs)

        ''' <summary>
        ''' Armazena o processo de mudança
        ''' </summary>
        ''' <remarks></remarks>
        Private _ChangedProcessEventArgs As ChangedProcessEventArgs

        ''' <summary>
        ''' Armazena o processo de exclusão
        ''' </summary>
        ''' <remarks></remarks>
        Private _ErasedProcessEventArgs As ErasedProcessEventArgs

        ''' <summary>
        ''' Armazena o processo de clonagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _ClonedProcessEventArgs As ClonedProcessEventArgs

        ''' <summary>
        ''' Armazena o processo de explosão
        ''' </summary>
        ''' <remarks></remarks>
        Private _ExplodedProcessEventArgs As ExplodedProcessEventArgs

        ''' <summary>
        ''' Monitoramento de comandos restritos
        ''' </summary>
        ''' <remarks></remarks>
        Private _RestrictProcessEventArgs As RestrictProcessEventArgs

        ''' <summary>
        ''' Armazena a transação corrente
        ''' </summary>
        ''' <remarks></remarks>
        Private _Transaction As Transaction

        ''' <summary>
        ''' Documento do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents _Document As Document

        ''' <summary>
        ''' Banco de dados do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents _Database As Database

        ''' <summary>
        ''' Editor do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents _Editor As Editor

        ''' <summary>
        ''' Determina se a reação esta ativa para entidade
        ''' </summary>
        ''' <remarks></remarks>
        Private _Active As Boolean

        ''' <summary>
        ''' Armazena a lista de comandos a serem monitorados pelo sistema
        ''' </summary>
        ''' <remarks></remarks>
        Private _SystemMonitoredCommands As List(Of String)

        ''' <summary>
        ''' Armazena os comandos de mudança
        ''' </summary>
        ''' <remarks></remarks>
        Private _ChangeCommands As New List(Of String)

        ''' <summary>
        ''' Armazena os comandos de clonagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _CloneCommands As New List(Of String)

        ''' <summary>
        ''' Armazena os comandos de explosão
        ''' </summary>
        ''' <remarks></remarks>
        Private _ExplodeCommands As New List(Of String)

        ''' <summary>
        ''' Armazena os comandos de exclusão
        ''' </summary>
        ''' <remarks></remarks>
        Private _EraseCommands As New List(Of String)

        ''' <summary>
        ''' Relação de comandos não tolerados para edição de itens controlados
        ''' </summary>
        ''' <remarks></remarks>
        Private _RestrictCommands As List(Of String)

        ''' <summary>
        ''' Relação de comandos monitorados pelo usuário
        ''' </summary>
        ''' <remarks></remarks>
        Private _UserMonitoredCommands As List(Of String)

        ''' <summary>
        ''' Armazena o nome do comando corrente
        ''' </summary>
        ''' <remarks></remarks>
        Private _GlobalCommandName As Object

        ''' <summary>
        ''' Determina se CommandWillStart esta ativo
        ''' </summary>
        ''' <remarks></remarks>
        Private _ActiveCommandWillStart As CommandEventHandler

        ''' <summary>
        '''  Determina se CommandCancelled esta ativo
        ''' </summary>
        ''' <remarks></remarks>
        Private _ActiveCommandCancelled As CommandEventHandler

        ''' <summary>
        '''  Determina se CommandEnded esta ativo
        ''' </summary>
        ''' <remarks></remarks>
        Private _ActiveCommandEnded As CommandEventHandler

        ''' <summary>
        ''' Classe complementar que detecta a pré chamada de comandos
        ''' </summary>
        Private WithEvents _VetoCommand As VetoCommand

        ''' <summary>
        ''' Variável reservada ao usuário
        ''' </summary>
        ''' <remarks></remarks>
        Private _Tag As Object

        ''' <summary>
        ''' Retorna os eventos de coleção de documentos
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DocumentCollection As DocumentCollection
            Get
                Return Me._DocumentCollection
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de todos os comandos monitorados pelo sistema
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SystemMonitoredCommands As List(Of String)
            Get
                Return Me._SystemMonitoredCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos de modificação monitorados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ChangeCommands As List(Of String)
            Get
                Return Me._ChangeCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos restritos monitorados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RestrictCommands As List(Of String)
            Get
                Return Me._RestrictCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos do usuário monitorados (Que não se enquadrem em nenhuma das demais categorias)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property UserMonitoredCommands As List(Of String)
            Get
                Return Me._UserMonitoredCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos de clonagem monitorados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CloneCommands As List(Of String)
            Get
                Return Me._CloneCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos de explosão monitorados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExplodeCommands As List(Of String)
            Get
                Return Me._ExplodeCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna a coleção de comandos de exclusão monitorados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property EraseCommands As List(Of String)
            Get
                Return Me._EraseCommands
            End Get
        End Property

        ''' <summary>
        ''' Retorna o processo de mudança corrente
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentChangedProcessEventArgs As ChangedProcessEventArgs
            Get
                Return Me._ChangedProcessEventArgs
            End Get
        End Property

        ''' <summary>
        ''' Retorna o processo de exclusão corrente
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentErasedProcessEventArgs As ErasedProcessEventArgs
            Get
                Return Me._ErasedProcessEventArgs
            End Get
        End Property

        ''' <summary>
        ''' Retorna o processo de clonagem corrente
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentClonedProcessEventArgs As ClonedProcessEventArgs
            Get
                Return Me._ClonedProcessEventArgs
            End Get
        End Property

        ''' <summary>
        ''' Retorna o processo de explosão corrente
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentExplodedProcessEventArgs As ExplodedProcessEventArgs
            Get
                Return Me._ExplodedProcessEventArgs
            End Get
        End Property

        ''' <summary>
        ''' Retorna o processo de controle de comandos não tolerados para edição de itens controlados
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentRestrictProcessEventArgs As RestrictProcessEventArgs
            Get
                Return Me._RestrictProcessEventArgs
            End Get
        End Property

        ''' <summary>
        ''' Documento do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Document As Document
            Get
                Return Me._Document
            End Get
        End Property

        ''' <summary>
        ''' Banco de dados do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Database As Database
            Get
                Return Me._Database
            End Get
        End Property

        ''' <summary>
        ''' Editor do AutoCad
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Editor As Editor
            Get
                Return Me._Editor
            End Get
        End Property

        ''' <summary>
        ''' Comando corrente 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property GlobalCommandName As Object
            Get
                Return Me._GlobalCommandName
            End Get
        End Property

        ''' <summary>
        ''' Transação
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Transaction As Transaction
            Get
                Return Me._Transaction
            End Get
        End Property

        ''' <summary>
        ''' Determina se CommandWillStart esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveCommandWillStart As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveCommandWillStart) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveCommandWillStart) = True Then
                        Me._ActiveCommandWillStart = AddressOf Reactors_Document_CommandWillStart
                        AddHandler Me._Document.CommandWillStart, Me._ActiveCommandWillStart
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveCommandWillStart) = False Then
                        RemoveHandler Me._Document.CommandWillStart, Me._ActiveCommandWillStart
                        Me._ActiveCommandWillStart = Nothing
                    End If

                End If

            End Set
        End Property

        ''' <summary>
        ''' Determina se CommandCancelled esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveCommandCancelled As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveCommandCancelled) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveCommandCancelled) = True Then
                        Me._ActiveCommandCancelled = AddressOf Reactors_Document_CommandCancelled
                        AddHandler Me._Document.CommandCancelled, Me._ActiveCommandCancelled
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveCommandCancelled) = False Then
                        RemoveHandler Me._Document.CommandCancelled, Me._ActiveCommandCancelled
                        Me._ActiveCommandCancelled = Nothing
                    End If

                End If

            End Set
        End Property

        ''' <summary>
        ''' Determina se CommandEnded esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveCommandEnded As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveCommandEnded) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveCommandEnded) = True Then
                        Me._ActiveCommandEnded = AddressOf Reactors_Document_CommandEnded
                        AddHandler Me._Document.CommandEnded, Me._ActiveCommandEnded
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveCommandEnded) = False Then
                        RemoveHandler Me._Document.CommandEnded, Me._ActiveCommandEnded
                        Me._ActiveCommandEnded = Nothing
                    End If

                End If

            End Set
        End Property

        ''' <summary>
        ''' Determina se o monitoramento da entidade esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Active As Boolean
            Get

                'Retorna o status de atividade
                Return Me._Active

            End Get
            Set(value As Boolean)

                'Compara o valor antigo e atual
                If value <> Me._Active Then

                    'Avalia value
                    If value = True Then

                        'Habilita a detecção de pré entrada de comandos
                        Me._VetoCommand.Active = True

                        'Habilita detecção de entrada de comandos
                        Me.ActiveCommandWillStart = True

                    Else

                        'Desabilita detecção de pré entrada de comandos
                        Me._VetoCommand.Active = False

                        'Desabilita detecção de entrada de comandos
                        Me.ActiveCommandWillStart = False

                    End If

                    'Desabilita detecção de cancelamento
                    Me.ActiveCommandCancelled = False

                    'Desabilita detecção de finalização
                    Me.ActiveCommandEnded = False

                    'Atualiza o status de atividade
                    Me._Active = value

                End If

            End Set
        End Property

        ''' <summary>
        ''' Variável reservada ao usuário
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
        ''' Limpa todos os processos
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Clean()

            'Destroi o monitor de eventos de mudança
            If IsNothing(Me._ChangedProcessEventArgs) = False Then
                Me._ChangedProcessEventArgs.Dispose()
                Me._ChangedProcessEventArgs = Nothing
            End If

            'Destroi o monitor de eventos de exclusão
            If IsNothing(Me._ErasedProcessEventArgs) = False Then
                Me._ErasedProcessEventArgs.Dispose()
                Me._ErasedProcessEventArgs = Nothing
            End If

            'Destroi o monitor de eventos de clonagem
            If IsNothing(Me._ClonedProcessEventArgs) = False Then
                Me._ClonedProcessEventArgs.Dispose()
                Me._ClonedProcessEventArgs = Nothing
            End If

            'Destroi o monitor de eventos de explosão
            If IsNothing(Me._ExplodedProcessEventArgs) = False Then
                Me._ExplodedProcessEventArgs.Dispose()
                Me._ExplodedProcessEventArgs = Nothing
            End If

            'Destroi o monitor de eventos para comandos restritos
            If IsNothing(Me._RestrictProcessEventArgs) = False Then
                Me._RestrictProcessEventArgs.Dispose()
                Me._RestrictProcessEventArgs = Nothing
            End If

            'Limpa a transação
            If IsNothing(Me._Transaction) = False Then
                If Me._Transaction.IsDisposed = False Then
                    Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                    Me._Transaction.Dispose()
                End If
                Me._Transaction = Nothing
            End If

            'Desabilita detecção de cancelamento
            Me.ActiveCommandCancelled = False

            'Desabilita detecção de finalização
            Me.ActiveCommandEnded = False

            'Limpa o comando global
            Me._GlobalCommandName = Nothing

        End Sub

        ''' <summary>
        ''' Detecta a entrada de comandos
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub Reactors_Document_CommandWillStart(sender As Object, e As CommandEventArgs)
            Try

                'Atualiza a interface com o AutoCad
                Me.Update(True)

                'Armazena o comando global
                Me._GlobalCommandName = e.GlobalCommandName

                'Avalia se o comando é monitorado (Usuário)
                If Me._UserMonitoredCommands.Contains(Me._GlobalCommandName) = True Then

                    'Dispara o evento de chamada para o comando monitorado pelo usupario
                    RaiseEvent MonitoredCommandStarted(Me, Me._GlobalCommandName)

                End If

                If Me._ChangeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de modificação

                    'Cria o monitor de eventos de mudança
                    Me._ChangedProcessEventArgs = New ChangedProcessEventArgs(Me)

                    'Habilita detecção de cancelamento
                    Me.ActiveCommandCancelled = True

                    'Habilita detecção de finalização
                    Me.ActiveCommandEnded = True

                ElseIf Me._EraseCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de exclusão

                    'Cria o monitor de eventos de exclusão
                    Me._ErasedProcessEventArgs = New ErasedProcessEventArgs(Me)

                    'Habilita detecção de cancelamento
                    Me.ActiveCommandCancelled = True

                    'Habilita detecção de finalização
                    Me.ActiveCommandEnded = True

                ElseIf Me._CloneCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de clonagem

                    'Cria o monitor de eventos de clonagem
                    Me._ClonedProcessEventArgs = New ClonedProcessEventArgs(Me)

                    'Habilita detecção de cancelamento
                    Me.ActiveCommandCancelled = True

                    'Habilita detecção de finalização
                    Me.ActiveCommandEnded = True

                ElseIf Me._ExplodeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de explosão

                    'Cria o monitor de eventos de explosão
                    Me._ExplodedProcessEventArgs = New ExplodedProcessEventArgs(Me)

                    'Habilita detecção de cancelamento
                    Me.ActiveCommandCancelled = True

                    'Habilita detecção de finalização
                    Me.ActiveCommandEnded = True

                ElseIf Me._RestrictCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é restrito

                    'Cria o monitor de eventos para comandos restritos
                    Me._RestrictProcessEventArgs = New RestrictProcessEventArgs(Me)

                    'Habilita detecção de cancelamento
                    Me.ActiveCommandCancelled = True

                    'Habilita detecção de finalização
                    Me.ActiveCommandEnded = True

                Else

                    'Limpa os processos
                    Me.Clean()

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try

        End Sub

        ''' <summary>
        '''  Detecta a finalização de comandos
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub Reactors_Document_CommandEnded(sender As Object, e As CommandEventArgs)

            'Declarações
            Dim ValidItem As Boolean = False

            Try

                'Avalia se o comando é monitorado (Usuário)
                If Me._UserMonitoredCommands.Contains(Me._GlobalCommandName) = True Then

                    'Dispara o evento de chamada para o comando monitorado pelo usupario
                    RaiseEvent MonitoredCommandEnded(Me, Me._GlobalCommandName)

                End If

                If Me._ChangeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de modificação

                    'Avalia as atividades do evento
                    If Me._ChangedProcessEventArgs.Count > 0 Then

                        'Seta o item como válido
                        ValidItem = True

                        'Constroi o evento de modificação
                        Me.InvokeChangedReactions(Me._ChangedProcessEventArgs)

                    End If

                ElseIf Me._EraseCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de exclusão

                    'Avalia as atividades do evento
                    If Me._ErasedProcessEventArgs.Count > 0 Then

                        'Seta o item como válido
                        ValidItem = True

                        'Constroi o evento de exclusão
                        Me.InvokeErasedReactions(Me._ErasedProcessEventArgs)

                    End If

                ElseIf Me._CloneCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de clonagem

                    'Otimiza o processo (Necessário quando no comando MIRROR a opção de excluir a entidade original é selecionada pelo usuário).
                    Me._ClonedProcessEventArgs.Optimize()

                    'Avalia as atividades do evento
                    If Me._ClonedProcessEventArgs.Count > 0 Then

                        'Seta o item como válido
                        ValidItem = True

                        'Constroi o evento de clonagem
                        Me.InvokeClonedReactions(Me._ClonedProcessEventArgs)

                    End If

                ElseIf Me._ExplodeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de explosão

                    'Avalia as atividades do evento
                    If Me._ExplodedProcessEventArgs.Count > 0 Then

                        'Seta o item como válido
                        ValidItem = True

                        'Constroi o evento de explosão
                        Me.InvokeExplodedReactions(Me._ExplodedProcessEventArgs)

                    End If

                ElseIf Me._RestrictCommands.Contains(Me._GlobalCommandName) = True Then

                    'Avalia as atividades do evento
                    If Me._RestrictProcessEventArgs.Count > 0 Then

                        'Seta o item como válido
                        ValidItem = True

                        'Constroi o evento de restrição
                        Me.InvokeRestrictReactions(Me._RestrictProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                If Me._ChangeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de modificação

                    'Avalia o nome do comando
                    If Me._GlobalCommandName = "GRIP_STRETCH" And ValidItem = True Then

                        'Remove a seleção de tela
                        Engine2.EntityInteration.Unselect()

                    End If

                End If

                'Limpa os processos
                Me.Clean()

            End Try

        End Sub

        ''' <summary>
        ''' Detecta o cancelamento de comandos
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub Reactors_Document_CommandCancelled(sender As Object, e As CommandEventArgs)
            Try

                'Avalia se o comando é monitorado (Usuário)
                If Me._UserMonitoredCommands.Contains(Me._GlobalCommandName) = True Then

                    'Dispara o evento de chamada para o comando monitorado pelo usupario
                    RaiseEvent MonitoredCommandCanceled(Me, Me._GlobalCommandName)

                End If

                If Me._ChangeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de modificação

                    'Avalia o nome do comando
                    If Me._GlobalCommandName = "GRIP_STRETCH" Then

                        'Remove a seleção de tela
                        Engine2.EntityInteration.Unselect()

                    End If

                    'Dispara o evento de cancelamento
                    RaiseEvent MonitoredCommandCanceled(Me, e.GlobalCommandName)

                ElseIf Me._EraseCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de exclusão

                    'Dispara o evento de cancelamento
                    RaiseEvent MonitoredCommandCanceled(Me, e.GlobalCommandName)

                ElseIf Me._CloneCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de clonagem

                    If Me._ClonedProcessEventArgs.Count > 0 Then

                        'Transfere o evento para finalização de comando
                        Reactors_Document_CommandEnded(sender, e)

                        'Finaliza a sub
                        Exit Sub

                    Else

                        'Dispara o evento de cancelamento
                        RaiseEvent MonitoredCommandCanceled(Me, e.GlobalCommandName)

                    End If

                ElseIf Me._ExplodeCommands.Contains(Me._GlobalCommandName) = True Then 'Avalia se o comando é de explosão

                    'Dispara o evento de cancelamento
                    RaiseEvent MonitoredCommandCanceled(Me, e.GlobalCommandName)

                Else

                    'Avalia se existe o monitor de eventos para comandos restritos
                    If IsNothing(Me._RestrictProcessEventArgs) = False Then

                        'Avalia as restrições
                        If Me._RestrictProcessEventArgs.Count > 0 Then

                            'Transfere o evento para finalização de comando
                            Reactors_Document_CommandEnded(sender, e)

                            'Finaliza a sub
                            Exit Sub

                        End If

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Limpa os processos
                Me.Clean()

            End Try

        End Sub

        ''' <summary>
        ''' Cria um novo reator (Não permite duplicações)
        ''' </summary>
        ''' <param name="Entity">Entidade a ser monitorada</param>
        ''' <param name="ReactorEvents">Coleção de eventos a serem monitorados</param>
        ''' <param name="Active">Determina se o monitoramento da entidade esta ativo</param>
        ''' <param name="Tag">Armazena valor do usuário</param>
        ''' <param name="PreventDuplication">Impede a entrada de duplicações</param>
        ''' <param name="Optimize">Determina se deve ser otimizado antes de incluir o novo reator</param>
        ''' <returns>Reactor</returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(Entity As Entity, Optional ReactorEvents As Engine2.Reactor.eReactorEvents = Reactor.eReactorEvents.ChangeReaction + Reactor.eReactorEvents.CloneReaction + Reactor.eReactorEvents.EraseReaction + Reactor.eReactorEvents.ExplodeReaction + Reactor.eReactorEvents.RestrictReaction, Optional Active As Boolean = True, Optional Tag As Object = Nothing, Optional PreventDuplication As Boolean = True, Optional Group As String = "", Optional Optimize As Boolean = False) As Reactor

            'Declarações
            Dim Reactor As Reactor = Nothing

            'Avalia a entidade
            If IsNothing(Entity) = False Then

                Try

                    'Verifica se deve previnir duplicações
                    If PreventDuplication = True Then

                        'Otimiza o sistema
                        Me.Optimize()

                        'Verifica se o reator existe
                        Reactor = MyBase.Find(Function(X As Reactor) X.Entity.Id.Equals(Entity.Id) = True)

                        'Avalia o reator
                        If IsNothing(Reactor) = True Then

                            'Cria o reator
                            Reactor = New Reactor(Me, Entity, ReactorEvents, Active, Tag, Group)

                            'Adiciona o Reactor2 na coleção
                            MyBase.Add(Reactor)

                        Else

                            'Seta o reator como nulo
                            Reactor = Nothing

                        End If

                    Else

                        'Avalia se deve otimizar
                        If Optimize = True Then

                            'Otimiza o sistema
                            Me.Optimize()

                        End If

                        'Cria o reator
                        Reactor = New Reactor(Me, Entity, ReactorEvents, Active, Tag, Group)

                        'Adiciona o Reactor2 na coleção
                        MyBase.Add(Reactor)

                    End If


                Catch ex As System.Exception

                    'Evento de erro
                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                    'Seta o reator como nulo
                    Reactor = Nothing

                Finally

                    'Avalia se o reator existe
                    If IsNothing(Reactor) = False Then

                        'Avalia se o reator pode disparar o evento Add
                        If Reactor.ReactorEvents.Contains("AddReactor") = True And Reactor.Active = True And Me.Active = True Then

                            'Dispara o evento
                            RaiseEvent ReactorAdded(Me, Reactor)

                        End If

                    End If

                End Try

            End If

            'Retorna o reator
            Return Reactor

        End Function

        ''' <summary>
        ''' Retorna o reator com base no id
        ''' </summary>
        ''' <param name="ObjectId"></param>
        ''' <returns></returns>
        Public Overloads Function Find(ObjectId As ObjectId) As Reactor

            'Declarações
            Find = Nothing

            Try

                'Verifica se existe
                Find = MyBase.Find(Function(X As Reactor) X.Entity.Id.Equals(ObjectId) = True)

                'Retorno
                Return Find

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try

        End Function

        ''' <summary>
        ''' Verifica se um reator existe com base no ID
        ''' </summary>
        ''' <param name="Objectid">Objectid</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(ObjectId As ObjectId) As Boolean
            Try

                'Verifica se existe
                If IsNothing(MyBase.Find(Function(X As Reactor) X.Entity.Id.Equals(ObjectId) = True)) = False Then
                    Return True
                Else
                    Return False
                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Function

        ''' <summary>
        ''' Remove um reator com base na entidade de controle
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <remarks></remarks>
        Public Overloads Sub Remove(ObjectId As ObjectId)
            Try

                'Obtem o reator
                Dim Reactor As Reactor = MyBase.Find(Function(X As Reactor) X.Entity.Id.Equals(ObjectId) = True)

                'Avalia o filtro
                If IsNothing(Reactor) = False Then

                    'Remove o reator
                    MyBase.Remove(Reactor)

                    'Avalia se o reator pode disparar o evento Removed
                    If Reactor.ReactorEvents.Contains("RemoveReactor") = True And Reactor.Active = True And Me.Active = True Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                Try

                                    'Atualiza a transação da classe
                                    Me._Transaction = Transaction

                                    'Dispara o evento de remoção
                                    RaiseEvent ReactorRemoved(Me, Reactor)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Atualiza dados da classe
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Update(Optional AutoCadInterfaceOnly As Boolean = False)
            Try

                'Verifica a opção do usuário
                If AutoCadInterfaceOnly = True Then

                    'Conecta com o AutoCad
                    Me._Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    If IsNothing(Me._Document) = False AndAlso Me._Document.IsActive = True Then
                        Me._Database = Me._Document.Database
                        Me._Editor = Me._Document.Editor
                    Else
                        Me._Database = Nothing
                        Me._Editor = Nothing
                    End If

                    'Atualiza a classe veto
                    If IsNothing(Me._VetoCommand) = False Then
                        Me._VetoCommand.Update()
                    End If

                Else

                    'Conecta com o AutoCad
                    Me._Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    If IsNothing(Me._Document) = False AndAlso Me._Document.IsActive = True Then
                        Me._Database = Me._Document.Database
                        Me._Editor = Me._Document.Editor
                    Else
                        Me._Database = Nothing
                        Me._Editor = Nothing
                    End If

                    'Atualiza a classe veto
                    If IsNothing(Me._VetoCommand) = False Then
                        Me._VetoCommand.Update()
                    End If

                    'Instancia a lista de comandos monitorados
                    With Me._SystemMonitoredCommands

                        'Limpa a coleção
                        .Clear()

                        'Carrega comandos de modificação
                        .AddRange(Me._ChangeCommands.ToArray)

                        'Carrega comandos de clonagem
                        .AddRange(Me._CloneCommands.ToArray)

                        'Carrega comandos de explosão
                        .AddRange(Me._ExplodeCommands.ToArray)

                        'Carrega comandos de exclusão
                        .AddRange(Me._EraseCommands.ToArray)

                        'Carrega comandos restritos
                        .AddRange(Me._RestrictCommands.ToArray)

                        'Carrega comandos monitorados do usuário
                        .AddRange(Me._UserMonitoredCommands.ToArray)

                        'Exclui duplicatas
                        .Distinct.ToList

                        'Ordena a coleção
                        .Sort()

                    End With

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Retorna os reatores válidos
        ''' </summary>
        ''' <param name="ActiveReactorOnly">Determina se somente reatores ativos serão retornados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValidReactors(Optional ActiveReactorOnly As Boolean = False) As List(Of Reactor)
            Try

                'Verifica o filtro do usuário
                If ActiveReactorOnly = True Then

                    'Retorna itens ativos
                    Return MyBase.FindAll(Function(X As Reactor) X.Active = True And X.Entity.IsErased = False)

                Else

                    'Retorna itens não ativos
                    Return MyBase.FindAll(Function(X As Reactor) X.Entity.IsErased = False)

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                'Retorno
                Return Nothing

            End Try
        End Function

        ''' <summary>
        ''' Filtro de exclusão de Optimize
        ''' Este filtro detecta e trata erro de processo corrompido.
        ''' </summary>
        ''' <param name="X"></param>
        ''' <returns></returns>
        <HandleProcessCorruptedStateExceptions>
        Private Function Filter(X As Reactor) As Boolean
            Try
                If IsNothing(X.Entity) = True Then
                    Return True
                ElseIf IsNothing(X.Entity.Id) = True Then
                    Return True
                ElseIf X.Entity.Id.IsNull = True Then
                    Return True
                ElseIf X.Entity.Id.IsValid = False Then
                    Return True
                ElseIf X.Entity.IsErased = True Then
                    Return True
                End If
                Return False
            Catch
                Return True
                Exit Try
            End Try
        End Function

        ''' <summary>
        ''' Otimiza a classe, excluindo reatores que tiveram suas entidades excluídas
        ''' </summary>
        Public Sub Optimize()
            Try

                'Remove itens inválidos
                MyBase.RemoveAll(Function(x As Reactor) Filter(x) = True)

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento ReactorsError
        ''' </summary>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Public Sub InvokeReactorsError(Message As String)
            Try

                'Constroi o evento de erro
                RaiseEvent ReactorsError(Me, Message)

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento ClonedReactions
        ''' </summary>
        ''' <param name="ClonedProcessEventArgs">ClonedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Sub InvokeClonedReactions(ByRef ClonedProcessEventArgs As ClonedProcessEventArgs)
            Try

                'Declarações
                'Dim SB As StringBuilder
                Dim Cancel As Boolean = False

                'Atualiza a variável de sistema
                Me._ClonedProcessEventArgs = ClonedProcessEventArgs

                'Paraliza a detecção de novos itens
                ClonedProcessEventArgs.StopEvents()

                'Evento de inicio de processo
                RaiseEvent StartedClonedReactions(Me, ClonedProcessEventArgs, Cancel)

                'Verifica se foi solicitado o cancelamento
                If Cancel = True Then

                    'Cancela a edição
                    Me.Undo()

                    ''Limpa StringBuilder
                    'SB = New StringBuilder

                    ''Monta o comando UNDO
                    'With SB
                    '    .Append("_.UNDO")
                    '    .Append(" ")
                    '    .Append("1")
                    '    .Append(" ")
                    'End With

                    ''Executa o comando UNDO
                    'Me._Document.SendStringToExecute(SB.ToString, True, False, False)

                Else

                    'Avalia o evento 
                    If ClonedProcessEventArgs.Count > 0 Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                'Passa a transação
                                Me._Transaction = Transaction

                                Try

                                    'Dispara o evento de restrição
                                    RaiseEvent ClonedReactions(Me, ClonedProcessEventArgs)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                        'Evento de fim de processo
                        RaiseEvent EndedClonedReactions(Me, ClonedProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Destroi a transação
                If IsNothing(Me._Transaction) = False Then
                    If Me._Transaction.IsDisposed = False Then
                        Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                        Me._Transaction.Dispose()
                    End If
                    Me._Transaction = Nothing
                End If

                'Destroi as informações do evento
                ClonedProcessEventArgs.Dispose()
                ClonedProcessEventArgs = Nothing
                Me._ClonedProcessEventArgs = Nothing

                'Atualiza os gráficos de tela
                Autodesk.AutoCAD.Internal.Utils.FlushGraphics()

                'Dispara o evento de fim de reação
                RaiseEvent EndReaction(Reactor.eReactorEvents.CloneReaction)

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento ErasedReactions
        ''' </summary>
        ''' <param name="ErasedProcessEventArgs">ErasedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Sub InvokeErasedReactions(ByRef ErasedProcessEventArgs As ErasedProcessEventArgs)
            Try

                'Declarações
                'Dim SB As StringBuilder
                Dim Cancel As Boolean = False

                'Atualiza a variável de sistema
                Me._ErasedProcessEventArgs = ErasedProcessEventArgs

                'Paraliza a detecção de novos itens
                ErasedProcessEventArgs.StopEvents()

                'Evento de inicio de processo
                RaiseEvent StartedErasedReactions(Me, ErasedProcessEventArgs, Cancel)

                'Verifica se foi solicitado o cancelamento
                If Cancel = True Then

                    'Cancela a edição
                    Me.Undo()

                    ''Limpa StringBuilder
                    'SB = New StringBuilder

                    ''Monta o comando UNDO
                    'With SB
                    '    .Append("_.UNDO")
                    '    .Append(" ")
                    '    .Append("1")
                    '    .Append(" ")
                    'End With

                    ''Executa o comando UNDO
                    'Me._Document.SendStringToExecute(SB.ToString, True, False, False)

                Else

                    'Avalia o evento
                    If ErasedProcessEventArgs.Count > 0 Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                'Passa a transação
                                Me._Transaction = Transaction

                                Try

                                    'Dispara o evento de restrição
                                    RaiseEvent ErasedReactions(Me, ErasedProcessEventArgs)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                        'Evento de fim de processo
                        RaiseEvent EndedErasedReactions(Me, ErasedProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Destroi a transação
                If IsNothing(Me._Transaction) = False Then
                    If Me._Transaction.IsDisposed = False Then
                        Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                        Me._Transaction.Dispose()
                    End If
                    Me._Transaction = Nothing
                End If

                'Destroi as informações do evento
                ErasedProcessEventArgs.Dispose()
                ErasedProcessEventArgs = Nothing
                Me._ErasedProcessEventArgs = Nothing

                'Atualiza os gráficos de tela
                Autodesk.AutoCAD.Internal.Utils.FlushGraphics()

                'Dispara o evento de fim de reação
                RaiseEvent EndReaction(Reactor.eReactorEvents.EraseReaction)

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento ExplodedReactions
        ''' </summary>
        ''' <param name="ExplodedProcessEventArgs">ExplodedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Sub InvokeExplodedReactions(ByRef ExplodedProcessEventArgs As ExplodedProcessEventArgs)
            Try

                'Declarações
                'Dim SB As StringBuilder
                Dim Cancel As Boolean = False

                'Atualiza a variável de sistema
                Me._ExplodedProcessEventArgs = ExplodedProcessEventArgs

                'Paraliza a detecção de novos itens
                ExplodedProcessEventArgs.StopEvents()

                'Evento de inicio de processo
                RaiseEvent StartedExplodedReactions(Me, ExplodedProcessEventArgs, Cancel)

                'Verifica se foi solicitado o cancelamento
                If Cancel = True Then

                    'Cancela a edição
                    Me.Undo()

                    ''Limpa StringBuilder
                    'SB = New StringBuilder

                    ''Monta o comando UNDO
                    'With SB
                    '    .Append("_.UNDO")
                    '    .Append(" ")
                    '    .Append("1")
                    '    .Append(" ")
                    'End With

                    ''Executa o comando UNDO
                    'Me._Document.SendStringToExecute(SB.ToString, True, False, False)

                Else

                    'Avalia o evento
                    If ExplodedProcessEventArgs.Count > 0 Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                'Passa a transação
                                Me._Transaction = Transaction

                                Try

                                    'Dispara o evento de restrição
                                    RaiseEvent ExplodedReactions(Me, ExplodedProcessEventArgs)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                        'Evento de fim de processo
                        RaiseEvent EndedExplodedReactions(Me, ExplodedProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Destroi a transação
                If IsNothing(Me._Transaction) = False Then
                    If Me._Transaction.IsDisposed = False Then
                        Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                        Me._Transaction.Dispose()
                    End If
                    Me._Transaction = Nothing
                End If

                'Destroi as informações do evento
                ExplodedProcessEventArgs.Dispose()
                ExplodedProcessEventArgs = Nothing
                Me._ExplodedProcessEventArgs = Nothing

                'Atualiza os gráficos de tela
                Autodesk.AutoCAD.Internal.Utils.FlushGraphics()

                'Dispara o evento de fim de reação
                RaiseEvent EndReaction(Reactor.eReactorEvents.ExplodeReaction)

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento ChangedReactions
        ''' </summary>
        ''' <param name="ChangedProcessEventArgs">ChangedProcessEventArgs</param>
        ''' <remarks></remarks>
        Public Sub InvokeChangedReactions(ByRef ChangedProcessEventArgs As ChangedProcessEventArgs)
            Try

                'Declarações
                'Dim SB As StringBuilder
                Dim Cancel As Boolean = False

                'Atualiza a variável de sistema
                Me._ChangedProcessEventArgs = ChangedProcessEventArgs

                'Paraliza a detecção de novos itens
                ChangedProcessEventArgs.StopEvents()

                'Evento de inicio de processo
                RaiseEvent StartedChangedReactions(Me, ChangedProcessEventArgs, Cancel)

                'Verifica se foi solicitado o cancelamento
                If Cancel = True Then

                    'Cancela a edição
                    Me.Undo()

                    ''Limpa StringBuilder
                    'SB = New StringBuilder

                    ''Monta o comando UNDO
                    'With SB
                    '    .Append("_.UNDO")
                    '    .Append(" ")
                    '    .Append("1")
                    '    .Append(" ")
                    'End With

                    ''Executa o comando UNDO
                    'Me._Document.SendStringToExecute(SB.ToString, True, False, False)

                Else

                    'Avalia o evento
                    If ChangedProcessEventArgs.Count > 0 Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                'Passa a transação
                                Me._Transaction = Transaction

                                Try

                                    'Dispara o evento de restrição
                                    RaiseEvent ChangedReactions(Me, ChangedProcessEventArgs)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                        'Evento de fim de processo
                        RaiseEvent EndedChangedReactions(Me, ChangedProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Destroi a transação
                If IsNothing(Me._Transaction) = False Then
                    If Me._Transaction.IsDisposed = False Then
                        Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                        Me._Transaction.Dispose()
                    End If
                    Me._Transaction = Nothing
                End If

                'Destroi as informações do evento
                ChangedProcessEventArgs.Dispose()
                ChangedProcessEventArgs = Nothing
                Me._ChangedProcessEventArgs = Nothing

                'Atualiza os gráficos de tela
                Autodesk.AutoCAD.Internal.Utils.FlushGraphics()

                'Dispara o evento de fim de reação
                RaiseEvent EndReaction(Reactor.eReactorEvents.ChangeReaction)

            End Try
        End Sub

        ''' <summary>
        ''' Chama o evento RestrictReactions
        ''' </summary>
        ''' <param name="RestrictProcessEventArgs">RestrictProcessEventArgsas</param>
        ''' <remarks></remarks>
        Public Sub InvokeRestrictReactions(ByRef RestrictProcessEventArgs As RestrictProcessEventArgs)
            Try

                'Declarações
                'Dim SB As StringBuilder
                Dim Cancel As Boolean = False

                'Atualiza a variável de sistema
                Me._RestrictProcessEventArgs = RestrictProcessEventArgs

                'Paraliza a detecção de novos itens
                RestrictProcessEventArgs.StopEvents()

                'Evento de inicio de processo
                RaiseEvent StartedRestrictReactions(Me, RestrictProcessEventArgs, Cancel)

                'Verifica se foi solicitado o cancelamento
                If Cancel = True Then

                    'Cancela a edição
                    Me.Undo()

                    ''Limpa StringBuilder
                    'SB = New StringBuilder

                    ''Monta o comando UNDO
                    'With SB
                    '    .Append("_.UNDO")
                    '    .Append(" ")
                    '    .Append("1")
                    '    .Append(" ")
                    'End With

                    ''Executa o comando UNDO
                    'Me._Document.SendStringToExecute(SB.ToString, True, False, False)

                Else

                    'Avalia o evento
                    If RestrictProcessEventArgs.Count > 0 Then

                        'Bloqueia o documento
                        Using Me._Editor.Document.LockDocument

                            'Abre a transação
                            Using Transaction As Transaction = Me._Database.TransactionManager.StartTransaction()

                                'Passa a transação
                                Me._Transaction = Transaction

                                Try

                                    'Dispara o evento de restrição
                                    RaiseEvent RestrictReactions(Me, RestrictProcessEventArgs)

                                    'Confirma a transação
                                    Transaction.Commit()

                                Catch ex As System.Exception

                                    'Aborta a transação
                                    Transaction.Abort()

                                    'Evento de erro
                                    RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

                                End Try

                            End Using

                        End Using

                        'Evento de fim de processo
                        RaiseEvent EndedRestrictReactions(Me, RestrictProcessEventArgs)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                RaiseEvent ReactorsError(Me, Engine2.Debug.GetExceptionMessage(ex))

            Finally

                'Destroi a transação
                If IsNothing(Me._Transaction) = False Then
                    If Me._Transaction.IsDisposed = False Then
                        Me._Transaction.TransactionManager.QueueForGraphicsFlush()
                        Me._Transaction.Dispose()
                    End If
                    Me._Transaction = Nothing
                End If

                'Destroi as informações do evento
                RestrictProcessEventArgs.Dispose()
                RestrictProcessEventArgs = Nothing
                Me._RestrictProcessEventArgs = Nothing

                'Atualiza os gráficos de tela
                Autodesk.AutoCAD.Internal.Utils.FlushGraphics()

                'Dispara o evento de fim de reação
                RaiseEvent EndReaction(Reactor.eReactorEvents.RestrictReaction)

            End Try
        End Sub

        ''' <summary>
        ''' Retorna a coleção dos Ids contidos no gerenciador de reatores
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetObjectIds(Match As System.Predicate(Of Reactor)) As ObjectIdCollection

            'Inicializa a coleção
            GetObjectIds = New ObjectIdCollection

            'Armazenamento
            Dim Reactors As List(Of Reactor)

            'Avalia se o filtro do usuário foi informado
            If IsNothing(Match) = False Then

                'Executa o filtro do usuário
                Reactors = MyBase.FindAll(Match)

            Else

                'Obtem toda a coleção
                Reactors = Me

            End If

            'Percorre a coleção
            For Index As Integer = 0 To Reactors.Count - 1

                'Adiciona a entidade
                GetObjectIds.Add(Reactors.Item(Index).ObjectId)

            Next

            'Retorno
            Return GetObjectIds

        End Function

        ''' <summary>
        ''' Retorna a coleção das entidades contidas no gerenciador de reatores
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEntitys(Match As System.Predicate(Of Reactor)) As DBObjectCollection

            'Inicializa a coleção
            GetEntitys = New DBObjectCollection

            'Armazenamento
            Dim Reactors As List(Of Reactor)

            'Avalia se o filtro do usuário foi informado
            If IsNothing(Match) = False Then

                'Executa o filtro do usuário
                Reactors = MyBase.FindAll(Match)

            Else

                'Obtem toda a coleção
                Reactors = Me

            End If

            'Percorre a coleção
            For Index As Integer = 0 To Reactors.Count - 1

                'Adiciona a entidade
                GetEntitys.Add(Reactors.Item(Index).Entity)

            Next

            'Retorno
            Return GetEntitys

        End Function

        ''' <summary>
        ''' Desfaz mudanças da última transação
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Undo()

            'Declarações
            Dim SB As StringBuilder

            'Limpa StringBuilder
            SB = New StringBuilder

            'Monta o comando UNDO
            With SB
                .Append("_.UNDO")
                .Append(" ")
                .Append("1")
                .Append(" ")
            End With

            'Executa o comando UNDO
            Me._Document.SendStringToExecute(SB.ToString, True, False, False)

        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

            'Limpa a transação
            Me._Transaction = Nothing

            'Limpa eventos de movimentação
            Me._ChangedProcessEventArgs = Nothing

            'Limpa eventos de exclusão
            Me._ErasedProcessEventArgs = Nothing

            'Limpa eventos de clonagem
            Me._ClonedProcessEventArgs = Nothing

            'Limpa eventos de explosão
            Me._ExplodedProcessEventArgs = Nothing

            'Limpa o comando corrente
            Me._GlobalCommandName = Nothing

            'Monitora comandos restritos
            Me._RestrictProcessEventArgs = Nothing

            'Desabilita CommandWillStart esta ativo
            Me._ActiveCommandWillStart = Nothing

            'Desabilita CommandCancelled
            Me._ActiveCommandCancelled = Nothing

            'Desabilita CommandEnded
            Me._ActiveCommandEnded = Nothing

            'Carrega os comandos de movimentação
            Me._ChangeCommands = New List(Of String)({"EXTEND", "GRIP_STRETCH", "MOVE", "STRETCH", "ROTATE"})

            'Carrega os comandos de clonagem
            Me._CloneCommands = New List(Of String)({"COPY", "MIRROR"})

            'Carrega os comandos de explosão
            Me._ExplodeCommands = New List(Of String)({"EXPLODE"})

            'Carrega os comandos de exclusão
            Me._EraseCommands = New List(Of String)({"ERASE"})

            'Carrega os comandos não permitidos
            Me._RestrictCommands = New List(Of String)({"FILLET", "CHAMFER", "PEDIT", "TRIM", "JOIN"})

            'Inicia a coleção de comandos monitorados
            Me._SystemMonitoredCommands = New List(Of String)

            'Inicia a coleção de comandos monitorados pelo usuário
            Me._UserMonitoredCommands = New List(Of String)

            'Carrega a lista de comandos monitorados
            Me.Update()

            'Instancia a classe de pré chamada de comandos
            Me._VetoCommand = New VetoCommand(Me, Me._SystemMonitoredCommands)

        End Sub

        ''' <summary>
        ''' Dispara o evento de pré chamada (Permite o cancelamento prévio do comando)
        ''' </summary>
        ''' <param name="VetoCommand"></param>
        ''' <param name="ObjectIds"></param>
        ''' <param name="Veto"></param>
        Private Sub _VetoCommand_InvokedVetoCommand(VetoCommand As VetoCommand, ObjectIds As List(Of ObjectId), ByRef Veto As Boolean) Handles _VetoCommand.InvokedVetoCommand
            Dim Cancel As Boolean = False
            Me._GlobalCommandName = VetoCommand.GlobalCommandName
            If Me._ChangeCommands.Contains(Me._GlobalCommandName) = True Then
                RaiseEvent StartingChangedReactions(Me, ObjectIds, Cancel)
                Veto = Cancel
            ElseIf Me._CloneCommands.Contains(Me._GlobalCommandName) = True Then
                RaiseEvent StartingClonedReactions(Me, ObjectIds, Cancel)
                Veto = Cancel
            ElseIf Me._ExplodeCommands.Contains(Me._GlobalCommandName) = True Then
                RaiseEvent StartingExplodedReactions(Me, ObjectIds, Cancel)
                Veto = Cancel
            ElseIf Me._EraseCommands.Contains(Me._GlobalCommandName) = True Then
                RaiseEvent StartingErasedReactions(Me, ObjectIds, Cancel)
                Veto = Cancel
            ElseIf Me._RestrictCommands.Contains(Me._GlobalCommandName) = True Then
                RaiseEvent StartingRestrictReactions(Me, ObjectIds, Cancel)
                Veto = Cancel
            End If
        End Sub

        ''' <summary>
        ''' Dispara o evento assim que um comando qualquer é chamado 
        ''' </summary>
        ''' <param name="VetoCommand"></param>
        ''' <param name="Veto"></param>
        Private Sub _VetoCommand_InvokedCommand(VetoCommand As VetoCommand, ByRef Veto As Boolean) Handles _VetoCommand.InvokedCommand
            Dim Cancel As Boolean = False
            Me._GlobalCommandName = VetoCommand.GlobalCommandName
            RaiseEvent CommandStarting(Me, Cancel)
            Veto = Cancel
        End Sub

        ''' <summary>
        ''' Detecta o documento que se torna corrente
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub AcadDocumentCollection_DocumentBecameCurrent(sender As Object, e As DocumentCollectionEventArgs) Handles _DocumentCollection.DocumentBecameCurrent
            If e.Document <> Me._Document Then
                Me.Update(True)
            End If
        End Sub

        ''' <summary>
        ''' Detecta o documento que está sendo fechado
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub _DocumentCollection_DocumentDestroyed(sender As Object, e As DocumentDestroyedEventArgs) Handles _DocumentCollection.DocumentDestroyed
            Me.Update(True)
            Me.Optimize()
        End Sub

    End Class

End Namespace




