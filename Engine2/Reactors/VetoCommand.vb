Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports System.Runtime.InteropServices

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2016
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    Public Class VetoCommand : Implements IDisposable

#Region "IDisposable Support"

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
                    Me.Active = False
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

        'Declarações
        Private _Reactors As Reactors
        Private WithEvents _Editor As Editor
        Private WithEvents _Document As Document
        Private WithEvents _DocumentCollection As DocumentCollection
        Private _ActiveLockModeChanged As DocumentLockModeChangedEventHandler
        Private _ActiveSelectionAdded As SelectionAddedEventHandler
        Private _ActiveCommandEnded As CommandEventHandler
        'Private _ActiveDocumentChanged As DocumentCollectionEventHandler
        Private _Commands As List(Of String)
        Private _GlobalCommandName As String

        ''' <summary>
        ''' Evento que detecta a seleção entidades do reator para um determinado comando
        ''' </summary>
        ''' <param name="ObjectIds">Coleção de id´s com entidades do reator selecionadas</param>
        ''' <param name="Veto"></param>
        Public Event InvokedVetoCommand(VetoCommand As VetoCommand, ObjectIds As List(Of ObjectId), ByRef Veto As Boolean)

        ''' <summary>
        ''' Evento que detecta a chamada de qualquer comando, mesmo fora da coleção informada.
        ''' Util para comandos de dificil controle como BREAK que não permite a detecção de sua coleção.
        ''' </summary>
        ''' <param name="VetoCommand"></param>
        ''' <param name="Veto"></param>
        Public Event InvokedCommand(VetoCommand As VetoCommand, ByRef Veto As Boolean)

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Reactors">Classe Reactors</param>
        ''' <param name="Commands">Comandos a serem monitorados</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors, Commands As List(Of String))
            Me._ActiveLockModeChanged = Nothing
            Me._ActiveSelectionAdded = Nothing
            Me._ActiveCommandEnded = Nothing
            'Me._ActiveDocumentChanged = Nothing
            Me._Reactors = Reactors
            Me._DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager 'Me._Reactors.DocumentCollection
            Me._Commands = Commands
            Me._Commands.Add("ERASE")
            Me._Commands.Add("EXPLODE")
            Me._Commands.Add("BREAK")
            Me.Update()
        End Sub

        ''' <summary>
        ''' Comando global
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property GlobalCommandName As String
            Get
                Return Me._GlobalCommandName
            End Get
        End Property

        ''' <summary>
        ''' Determina se a classe está ou não ativa
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Active As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveLockModeChanged) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveLockModeChanged) = True Then
                        Me._ActiveLockModeChanged = AddressOf Me._DocumentCollection_DocumentLockModeChanged
                        AddHandler Me._DocumentCollection.DocumentLockModeChanged, Me._ActiveLockModeChanged
                    End If

                    ''Habilita
                    'If IsNothing(Me._ActiveDocumentChanged) = True Then
                    '    Me._ActiveDocumentChanged = AddressOf Me._DocumentChanged
                    '    AddHandler Me._DocumentCollection.DocumentActivated, Me._ActiveDocumentChanged
                    '    AddHandler Me._DocumentCollection.DocumentBecameCurrent, Me._ActiveDocumentChanged
                    '    AddHandler Me._DocumentCollection.DocumentCreated, Me._ActiveDocumentChanged
                    '    AddHandler Me._DocumentCollection.DocumentCreationCanceled, Me._ActiveDocumentChanged
                    '    AddHandler Me._DocumentCollection.DocumentCreateStarted, Me._ActiveDocumentChanged
                    'End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveLockModeChanged) = False Then
                        RemoveHandler Me._DocumentCollection.DocumentLockModeChanged, Me._ActiveLockModeChanged
                        Me._ActiveLockModeChanged = Nothing
                    End If

                    ''Habilita
                    'If IsNothing(Me._ActiveDocumentChanged) = True Then
                    '    Me._ActiveDocumentChanged = AddressOf Me._DocumentChanged
                    '    RemoveHandler Me._DocumentCollection.DocumentActivated, Me._ActiveDocumentChanged
                    '    RemoveHandler Me._DocumentCollection.DocumentBecameCurrent, Me._ActiveDocumentChanged
                    '    RemoveHandler Me._DocumentCollection.DocumentCreated, Me._ActiveDocumentChanged
                    '    RemoveHandler Me._DocumentCollection.DocumentCreationCanceled, Me._ActiveDocumentChanged
                    '    RemoveHandler Me._DocumentCollection.DocumentCreateStarted, Me._ActiveDocumentChanged
                    '    Me._ActiveDocumentChanged = Nothing
                    'End If

                    'Desabilita
                    Me.ActiveSelectionAdded = False

                    'Desabilita
                    Me.ActiveCommandEnded = False

                End If

            End Set

        End Property

        ''' <summary>
        ''' Determina se a detecção de entidades está ativa
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveSelectionAdded As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveSelectionAdded) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveSelectionAdded) = True Then
                        Me._ActiveSelectionAdded = AddressOf Me._Editor_SelectionAdded
                        AddHandler Me._Editor.SelectionAdded, Me._ActiveSelectionAdded
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveSelectionAdded) = False Then
                        RemoveHandler Me._Editor.SelectionAdded, Me._ActiveSelectionAdded
                        Me._ActiveSelectionAdded = Nothing
                    End If

                End If

            End Set

        End Property

        ''' <summary>
        ''' Determina se a detecção de finalização de comando estará ativa
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
                        Me._ActiveCommandEnded = AddressOf Me._DocumentCommandEnded
                        AddHandler Me._Document.CommandEnded, Me._ActiveCommandEnded
                        AddHandler Me._Document.CommandCancelled, Me._ActiveCommandEnded
                        AddHandler Me._Document.CommandFailed, Me._ActiveCommandEnded
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveCommandEnded) = False Then
                        RemoveHandler Me._Document.CommandEnded, Me._ActiveCommandEnded
                        RemoveHandler Me._Document.CommandCancelled, Me._ActiveCommandEnded
                        RemoveHandler Me._Document.CommandFailed, Me._ActiveCommandEnded
                        Me._ActiveCommandEnded = Nothing
                    End If

                End If

            End Set

        End Property

        ''' <summary>
        ''' Atualiza dados da classe
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Update()

            'Conecta com o AutoCad
            Me._Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If IsNothing(Me._Document) = False AndAlso Me._Document.IsActive = True Then
                Me._Editor = Me._Document.Editor
            Else
                Me._Editor = Nothing
            End If


        End Sub

        '''' <summary>
        '''' Detecta a mudança de documento e atualiza o status
        '''' </summary>
        '''' <param name="sender"></param>
        '''' <param name="e"></param>
        'Private Sub _DocumentChanged(sender As Object, e As DocumentCollectionEventArgs)
        '    Me.Update()
        'End Sub

        ''' <summary>
        ''' Dispara o evento para seleções pré comando
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub _DocumentCollection_DocumentLockModeChanged(sender As Object, e As DocumentLockModeChangedEventArgs)
            If IsNothing(Me._Editor) = False Then
                '==========================================================================================
                'Este try resolve a falha que ocorre na abertura de arquivos em Me._Editor.SelectImplied()
                'Não parece ser necessário nesta versão, mantido apenas por questão de compatibilidade 2009
                Try
                    Me._Editor.SelectImplied()
                Catch
                    Exit Sub
                    Exit Try
                End Try
                '==========================================================================================
                Dim Veto As Boolean = False
                Dim ObjectIds As New List(Of ObjectId)
                Try
                    Me._GlobalCommandName = e.GlobalCommandName
                    If Me._GlobalCommandName <> "" Then
                        If Me._GlobalCommandName.Contains("#") = False Then
                            RaiseEvent InvokedCommand(Me, Veto)
                            If Veto = False Then
                                If Me._Commands.Contains(Me._GlobalCommandName) = True And IsNothing(Me._Editor.SelectImplied.Value) = False Then
                                    For Each Id As ObjectId In Me._Editor.SelectImplied.Value.GetObjectIds
                                        If Me._Reactors.Contains(Id) = True Then
                                            ObjectIds.Add(Id)
                                        End If
                                    Next
                                    If ObjectIds.Count > 0 Then
                                        RaiseEvent InvokedVetoCommand(Me, ObjectIds, Veto)
                                        If Veto = True Then
                                            e.Veto()
                                            Engine2.EntityInteration.Unselect()
                                            Me._Editor.WriteMessage(vbCr & "*Cancel*" & vbCr)
                                        End If
                                        Me.CancelMonitoring()
                                    End If
                                Else
                                    Me.ActiveSelectionAdded = True
                                    Me.ActiveCommandEnded = True
                                End If
                            Else
                                If Veto = True Then
                                    e.Veto()
                                    Engine2.EntityInteration.Unselect()
                                    Me._Editor.WriteMessage(vbCr & "*Cancel*" & vbCr)
                                End If
                                Me.CancelMonitoring()
                            End If
                        End If
                    End If
                Catch ex As System.Exception
                    Me.CancelMonitoring()
                    Exit Try
                    'Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Detecta a seleção de entidades
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub _Editor_SelectionAdded(sender As Object, e As SelectionAddedEventArgs)
            If IsNothing(Me._Editor) = False Then
                Dim Veto As Boolean = False
                Dim ObjectIds As New List(Of ObjectId)
                Try
                    If Me._GlobalCommandName <> "" Then
                        If Me._GlobalCommandName.Contains("#") = False Then
                            If Me._Commands.Contains(Me._GlobalCommandName) = True Then
                                For Each Id In e.AddedObjects.GetObjectIds
                                    If Me._Reactors.Contains(Id) = True Then
                                        ObjectIds.Add(Id)
                                    End If
                                Next
                                If ObjectIds.Count > 0 Then
                                    RaiseEvent InvokedVetoCommand(Me, ObjectIds, Veto)
                                    If Veto = True Then
                                        Engine2.AcadInterface.CancelCommand(True)
                                        Engine2.EntityInteration.Unselect()
                                    End If
                                End If
                            End If
                        Else
                            Me.CancelMonitoring()
                        End If
                    Else
                        Me.CancelMonitoring()
                    End If
                Catch ex As Exception
                    Me.CancelMonitoring()
                    Exit Try
                    'Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Detecta a finalização do comando
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub _DocumentCommandEnded(sender As Object, e As CommandEventArgs)
            Me.CancelMonitoring()
        End Sub

        ''' <summary>
        ''' Cancela o monitoramento
        ''' </summary>
        Private Sub CancelMonitoring()
            If IsNothing(Me._Editor) = False Then
                Me.ActiveSelectionAdded = False
                Me.ActiveCommandEnded = False
            End If
        End Sub

    End Class

End Namespace

