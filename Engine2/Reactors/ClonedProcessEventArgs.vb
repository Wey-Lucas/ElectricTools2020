Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2016
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    '''  Armazena a coleção de itens relacionados a clonagem de entidades
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class ClonedProcessEventArgs : Inherits List(Of ClonedProcess) : Implements IDisposable

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
                    ActiveSelectionAdded = False
                    ActiveSelectionRemoved = False
                    ActiveObjectModified = False
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
        ''' Controla se o evento SelectionAdded está ativo
        ''' </summary>
        Private _ActiveSelectionAdded As SelectionAddedEventHandler

        ''' <summary>
        ''' Controla se o evento SelectionRemoved está ativo
        ''' </summary>
        Private _ActiveSelectionRemoved As SelectionRemovedEventHandler

        ''' <summary>
        ''' Controla se o evento ObjectModified está ativo
        ''' </summary>
        Private _ActiveObjectModified As ObjectEventHandler

        ''' <summary>
        ''' Assinatura da Xdata temporária
        ''' </summary>
        ''' <remarks></remarks>
        Const Assign As String = "ClonedProcess"

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
        ''' Classe superior
        ''' </summary>
        ''' <remarks></remarks>
        Private _Reactors As Reactors

        ''' <summary>
        ''' Determina se o evento foi fabricado pelo usuário
        ''' </summary>
        ''' <remarks></remarks>
        Private _IsMakeEvent As Boolean

        ''' <summary>
        ''' Armazena valor do usuário
        ''' </summary>
        ''' <remarks></remarks>
        Private _Tag As Object

        ''' <summary>
        ''' Determina se SelectionAdded esta ativo
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
        ''' Determina se SelectionAdded esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveSelectionRemoved As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveSelectionRemoved) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveSelectionRemoved) = True Then
                        Me._ActiveSelectionRemoved = AddressOf Me._Editor_SelectionRemoved
                        AddHandler Me._Editor.SelectionRemoved, Me._ActiveSelectionRemoved
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveSelectionRemoved) = False Then
                        RemoveHandler Me._Editor.SelectionRemoved, Me._ActiveSelectionRemoved
                        Me._ActiveSelectionRemoved = Nothing
                    End If

                End If

            End Set
        End Property

        ''' <summary>
        ''' Determina se ObjectModified esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property ActiveObjectModified As Boolean
            Get

                'Retorna o Status atual
                Return If(IsNothing(Me._ActiveObjectModified) = True, False, True)

            End Get
            Set(value As Boolean)

                'Avalia o valor informado
                If value = True Then

                    'Habilita
                    If IsNothing(Me._ActiveObjectModified) = True Then
                        Me._ActiveObjectModified = AddressOf Me._Database_ObjectModified
                        AddHandler Me._Database.ObjectModified, Me._ActiveObjectModified
                    End If

                Else

                    'Desabilita
                    If IsNothing(Me._ActiveObjectModified) = False Then
                        RemoveHandler Me._Database.ObjectModified, Me._ActiveObjectModified
                        Me._ActiveObjectModified = Nothing
                    End If

                End If

            End Set
        End Property

        ''' <summary>
        ''' Valor do usuário (Para uso de MakeEvent)
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
        ''' Determina se o evento foi fabricado pelo usuário
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsMakeEvent As Boolean
            Get
                Return Me._IsMakeEvent
            End Get
        End Property

        ''' <summary>
        ''' Comando global
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property GlobalCommandName As Object
            Get
                Return Me._Reactors.GlobalCommandName
            End Get
        End Property

        ''' <summary>
        ''' Transação corrente
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Transaction As Transaction
            Get
                Return Me._Reactors.Transaction
            End Get
        End Property

        ''' <summary>
        ''' Exclui itens inválidos do processo. Considera processos onde não existem itens clonados.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Optimize()
            Try

                'Exclui itens inválidos
                MyBase.RemoveAll(Function(X As ClonedProcess) X.ClonedItems.Count = 0)

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Detecta a seleção (pós comando)
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _Editor_SelectionAdded(sender As Object, e As SelectionAddedEventArgs)
            Try

                'Inclui a seleção
                Me.AddObject(e.AddedObjects.GetObjectIds)

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Detecta a exclusão da seleção (pós comando)
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _Editor_SelectionRemoved(sender As Object, e As SelectionRemovedEventArgs)
            Try

                'Remove a seleção
                Me.RemoveObject(e.RemovedObjects.GetObjectIds)

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Detecta a confirmação de edição 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _Database_ObjectModified(sender As Object, e As ObjectEventArgs)
            Try

                'Avalia se o item é entidade
                If Engine2.EntityInteration.IsEntity(e.DBObject) = True Then

                    'Armazena o processo
                    Dim ClonedProcess As ClonedProcess

                    'Armazena o handle original
                    Dim OriginalHandle As String

                    'Verifica se o item é assinado
                    If Engine2.XDataEngine.Contains(e.DBObject, Assign) = True Then

                        'Obtem o handle original
                        OriginalHandle = Engine2.XDataEngine.GetXData(e.DBObject, Assign)

                        'Verifica se e o item original
                        If OriginalHandle <> e.DBObject.Handle.ToString Then

                            'Obtem o processo
                            ClonedProcess = MyBase.Find(Function(X As ClonedProcess) X.OriginalState.Handle.ToString.Equals(OriginalHandle) = True)

                            'Verifica se o processo existe
                            If IsNothing(ClonedProcess) = False Then

                                'Verifica se o item já existe na coleção
                                If IsNothing(ClonedProcess.ClonedItems.Find(Function(X As DBObject) X.ObjectId.Equals(e.DBObject.ObjectId) = True)) = True Then

                                    'Adiciona o item na coleção
                                    ClonedProcess.ClonedItems.Add(e.DBObject)

                                End If

                            End If

                        End If

                        'Apaga a assinatura
                        Engine2.XDataEngine.DeleteXdata(e.DBObject, Assign)

                    End If

                End If

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try

        End Sub

        ''' <summary>
        ''' Adiciona item
        ''' </summary>
        ''' <param name="ObjectIds">Array</param>
        ''' <remarks></remarks>
        Private Sub AddObject(ObjectIds() As ObjectId)
            Try

                'Percorre a coleção
                For Each ObjectId As ObjectId In ObjectIds

                    'Verifica a validade do item
                    If Me.ValidId(ObjectId) = True Then

                        'Verifica se o item existe
                        If IsNothing(MyBase.Find(Function(X As ClonedProcess) X.ObjectId.Equals(ObjectId) = True)) = True Then

                            'Grava a Xdata
                            Engine2.XDataEngine.SetXData(ObjectId, Assign, ObjectId.Handle.ToString)

                            'Adiciona o item
                            Me.Add(New ClonedProcess(Me._Reactors, ObjectId))

                        End If

                    End If

                Next

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub


        ''' <summary>
        ''' Remove itens com base na coleção Id
        ''' </summary>
        ''' <param name="ObjectIds"></param>
        ''' <remarks></remarks>
        Private Sub RemoveObject(ObjectIds() As ObjectId)
            Try

                'Converte a coleção para ArrayList
                Dim ArrayList As New ArrayList(ObjectIds)

                'Remove os itens
                Me.RemoveAll(Function(X As ClonedProcess) ArrayList.Contains(X.ObjectId) = True)

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Valida o Id de acordo com o evento programado
        ''' </summary>
        ''' <param name="ObjectId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ValidId(ObjectId As ObjectId) As Boolean
            Try

                'Verifiva se existe e se é válido
                If IsNothing(Me._Reactors.Find(Function(X As Reactor) X.ObjectId.Equals(ObjectId) = True And X.Active = True And X.ReactorEvents.Contains("CloneReaction") = True)) = False Then

                    'Indica que é válido
                    Return True

                Else

                    'Indica que é falso
                    Return False

                End If

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Function

        ''' <summary>
        ''' Paraliza os eventos de detecção
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StopEvents()
            ActiveSelectionAdded = False
            ActiveSelectionRemoved = False
            ActiveObjectModified = False
        End Sub

        ''' <summary>
        ''' Adicionar processo
        ''' </summary>
        ''' <param name="OriginalState">Entidade em seu estado original</param>
        ''' <param name="ClonedItems">Coleção de entidades resultantes da clonagem</param>
        ''' <param name="PreventDuplication">Impede a entrada de duplicações</param>
        ''' <remarks></remarks>
        Public Overloads Sub Add(OriginalState As Entity, ClonedItems As List(Of Entity), Optional PreventDuplication As Boolean = True)
            Try

                'Avalia o filtro
                If PreventDuplication = True Then

                    'Verifica se existe
                    If Me.Contains(OriginalState) = False Then

                        'Adiciona
                        MyBase.Add(New ClonedProcess(Me._Reactors, OriginalState, ClonedItems))

                    End If

                Else

                    'Adiciona
                    MyBase.Add(New ClonedProcess(Me._Reactors, OriginalState, ClonedItems))

                End If

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

            End Try
        End Sub

        ''' <summary>
        ''' Retorna a coleção dos Ids
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOriginalStateObjectIds() As ObjectIdCollection
            Try

                'Inicializa a coleção
                GetOriginalStateObjectIds = New ObjectIdCollection

                'Percorre a coleção
                For Index As Integer = 0 To MyBase.Count - 1

                    'Adiciona a entidade
                    GetOriginalStateObjectIds.Add(MyBase.Item(Index).ObjectId)

                Next

                'Retorno
                Return GetOriginalStateObjectIds

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

                'Retorno
                Return Nothing

            End Try
        End Function

        ''' <summary>
        ''' Retorna a coleção das entidades
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOriginalStateEntitys() As DBObjectCollection
            Try

                'Declarações
                GetOriginalStateEntitys = New DBObjectCollection

                'Percorre a coleção
                For Index As Integer = 0 To MyBase.Count - 1

                    'Adiciona a entidade
                    GetOriginalStateEntitys.Add(MyBase.Item(Index).OriginalState)

                Next

                'Retorno
                Return GetOriginalStateEntitys

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

                'Retorno
                Return Nothing

            End Try
        End Function

        ''' <summary>
        ''' Determina se todas as seleções são do mesmo tipo
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsTypeEquals(Optional Match As System.Predicate(Of ClonedProcess) = Nothing) As Boolean
            Try

                'Cria a busca e compara as quantidades
                If MyBase.FindAll(Match).Count <> MyBase.Count Then

                    'Retorno
                    Return False

                Else

                    'Retorno
                    Return True

                End If

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

                'Retorno
                Return False

            End Try
        End Function

        ''' <summary>
        ''' Determina se o processo já existe
        ''' </summary>
        ''' <param name="OriginalState">Estado original</param>
        ''' <param name="ClonedItems">Itens clonados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(OriginalState As Entity, Optional ClonedItems As List(Of Entity) = Nothing) As Boolean
            Try

                'Declarações
                Dim Process As ClonedProcess

                'Obtem o processo
                Process = MyBase.Find(Function(X As ClonedProcess) X.OriginalState.Equals(OriginalState) = True)

                'Avalia NewState
                If IsNothing(ClonedItems) = True Then

                    'Verifica se existe
                    If IsNothing(Process) = True Then

                        'Retorno
                        Return False

                    Else

                        'Retorno
                        Return True

                    End If

                Else

                    'Avalia as quantidades
                    If ClonedItems.Count = Process.ClonedItems.Count Then

                        'Percorre a coleção
                        For Each Entity As Entity In ClonedItems

                            'Avalia se existe
                            If Process.ClonedItems.Contains(Entity) = False Then

                                'Retorno
                                Return False

                            End If

                        Next

                    Else

                        'Retorno
                        Return False

                    End If

                    'Retorno
                    Return True

                End If

            Catch ex As System.Exception

                'Evento de erro
                Me._Reactors.InvokeReactorsError(Engine2.Debug.GetExceptionMessage(ex))

                'Retorno
                Return False

            End Try
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Reactors">Classe proprietária</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors)

            'Armazena reactors
            Me._Reactors = Reactors

            'Atualiza documento
            Me._Document = Me._Reactors.Document

            'Atualiza database
            Me._Database = Me._Reactors.Database

            'Atualiza editor
            Me._Editor = Me._Reactors.Editor

            'Seta o status inicial de SelectionAdded
            Me._ActiveSelectionAdded = Nothing

            'Seta o status inicial de SelectionRemoved
            Me._ActiveSelectionRemoved = Nothing

            'Seta o status inicial de ObjectModified
            Me._ActiveObjectModified = Nothing

            'Obtem a pré-seleção
            If Me._Editor.SelectImplied.Status = PromptStatus.OK Then
                Me.AddObject(Me._Editor.SelectImplied.Value.GetObjectIds)
            End If

            'Habilita os eventos de monitoramento
            ActiveSelectionAdded = True
            ActiveSelectionRemoved = True
            ActiveObjectModified = True

            'Informa o evento não fabricado
            Me._IsMakeEvent = False

            'Seta o valor do usuário
            Me._Tag = Nothing

        End Sub

    End Class

    ''' <summary>
    ''' Armazena as informações de evolução de entidades editadas no AutoCAD
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClonedProcess : Inherits ProcessBase

        ''' <summary>
        ''' Coleção de clones
        ''' </summary>
        ''' <remarks></remarks>
        Private _ClonedItems As List(Of Entity)

        ''' <summary>
        ''' Clones
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ClonedItems As List(Of Entity)
            Get
                Return Me._ClonedItems
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Reactors">Classe proprietária</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors, ObjectId As ObjectId)

            'Carrega a classe base
            MyBase.New(Reactors, ObjectId)

            'Inicia a coleção
            Me._ClonedItems = New List(Of Entity)

        End Sub

        ''' <summary>
        ''' Construtor (Para uso de MakeEvent)
        ''' </summary>
        ''' <param name="Reactors">Classe proprietária</param>
        ''' <param name="OriginalState">Entidade em seu estado original</param>
        ''' <param name="ClonedItems">Coleção de entidades resultantes da clonagem</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors, OriginalState As Entity, ClonedItems As List(Of Entity))

            'Carrega a classe base
            MyBase.New(Reactors, OriginalState.ObjectId)

            'Carrega a coleção
            Me._ClonedItems = ClonedItems

        End Sub

    End Class

End Namespace