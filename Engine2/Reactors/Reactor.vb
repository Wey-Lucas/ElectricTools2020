Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Numerics

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2016
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Reator item de reactors
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Reactor

        ''' <summary>
        ''' Classe proprietária
        ''' </summary>
        ''' <remarks></remarks>
        Private _Reactors As Reactors

        ''' <summary>
        ''' Permite a soma de itens de um Enumerador para fim de múltiplas seleções
        ''' </summary>
        ''' <remarks></remarks>
        Private _GeometricSequence As Engine2.GeometricSequence

        ''' <summary>
        ''' Soma dos eventos relacionados ao objeto 
        ''' </summary>
        ''' <remarks></remarks>
        Private _NumericSequence As BigInteger

        ''' <summary>
        ''' Entidade monitorada
        ''' </summary>
        ''' <remarks></remarks>
        Private _Entity As Entity

        ''' <summary>
        ''' Determina se a reação esta ativa para entidade
        ''' </summary>
        ''' <remarks></remarks>
        Private _Active As Boolean

        ''' <summary>
        ''' Armazena valor do usuário
        ''' </summary>
        ''' <remarks></remarks>
        Private _Tag As Object

        ''' <summary>
        ''' Handle
        ''' </summary>
        ''' <remarks></remarks>
        Private _Handle As Handle

        ''' <summary>
        ''' ObjectId
        ''' </summary>
        ''' <remarks></remarks>
        Private _ObjectId As ObjectId

        ''' <summary>
        ''' Grupo
        ''' </summary>
        ''' <remarks></remarks>
        Private _Group As String

        ''' <summary>
        ''' Enumera os eventos associados a entidade a ser monitorada
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eReactorEvents

            ''' <summary>
            ''' Captura a inclusão de um reator
            ''' </summary>
            ''' <remarks></remarks>
            AddReactor = 1

            ''' <summary>
            ''' Captura a exclusão de um reator
            ''' </summary>
            ''' <remarks></remarks>
            RemoveReactor = 2

            ''' <summary>
            ''' Captura evento de clonagem (Comandos COPY e MIRROR)
            ''' </summary>
            ''' <remarks></remarks>
            CloneReaction = 4

            ''' <summary>
            ''' Captura evento de exclusão (Comando ERASE)
            ''' </summary>
            ''' <remarks></remarks>
            EraseReaction = 8

            ''' <summary>
            ''' Captura evento de explosão (Comando EXPLODE)
            ''' </summary>
            ''' <remarks></remarks>
            ExplodeReaction = 16

            ''' <summary>
            ''' Captura evento de mudança (Comandos EXTEND, GRIP_STRETCH, MOVE e STRETCH)
            ''' </summary>
            ''' <remarks></remarks>
            ChangeReaction = 32

            ''' <summary>
            ''' Captura quando a entidade é editada por comandos restritos
            ''' </summary>
            ''' <remarks></remarks>
            RestrictReaction = 64

        End Enum

        ''' <summary>
        ''' Classe proprietária
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Reactors As Reactors
            Get
                Return Me._Reactors
            End Get
        End Property

        ''' <summary>
        ''' Lista de eventos associados
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ReactorEvents As ArrayList
            Get
                Return Me._GeometricSequence.GetItemsInSequence(Me._NumericSequence)
            End Get
        End Property

        ''' <summary>
        ''' Entidade monitorada
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Entity As Entity
            Get
                Return Me._Entity
            End Get
        End Property

        ''' <summary>
        ''' Retorna o handle
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Handle As Handle
            Get
                Return Me._Handle
            End Get
        End Property

        ''' <summary>
        ''' Retorna o ObjectId
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ObjectId As ObjectId
            Get
                Return Me._ObjectId
            End Get
        End Property

        ''' <summary>
        ''' Armazena valor livre para uso do usuário
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
        ''' Determina se o monitoramento da entidade esta ativo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Active As Object
            Get
                Return Me._Active
            End Get
            Set(value As Object)
                Me._Active = value
            End Set
        End Property

        ''' <summary>
        ''' Grupo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Group As String
            Get
                Return Me._Group
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Reactors">Classe proprietária</param>
        ''' <param name="Entity">Entidade a ser monitorada</param>
        ''' <param name="ReactorEvents">Coleção de eventos a serem monitorados</param>
        ''' <param name="Active">Determina se o monitoramento da entidade esta ativo</param>
        ''' <param name="Tag">Armazena valor do usuário</param>
        ''' <param name="Group">Grupo</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors, Entity As Entity, Optional ReactorEvents As Engine2.Reactor.eReactorEvents = Reactor.eReactorEvents.ChangeReaction + Reactor.eReactorEvents.CloneReaction + Reactor.eReactorEvents.EraseReaction + Reactor.eReactorEvents.ExplodeReaction + Reactor.eReactorEvents.RestrictReaction, Optional Active As Boolean = True, Optional Tag As Object = Nothing, Optional Group As String = "")

            'Atualiza a classe proprietária
            Me._Reactors = Reactors

            'Inicia a classe GeometricSequence
            Me._GeometricSequence = New Engine2.GeometricSequence

            'Cria as sequencias válidas com base no enumerador 'eReactorEvents'
            Me._GeometricSequence.Sequences.Add(GetType(eReactorEvents))

            'Armazena a entidade monitorada
            Me._Entity = Entity

            'Armazena ObjectId
            Me._ObjectId = Me._Entity.ObjectId

            'Armazena o handle
            Me._Handle = Me._Entity.Handle

            'Armazena o número das reações
            Me._NumericSequence = CType(ReactorEvents, BigInteger)

            'Determina se a reação esta ativa para entidade
            Me._Active = Active

            'Armazena valor do usuário
            Me._Tag = Tag

            'Armazena o grupo
            Me._Group = Group

        End Sub


    End Class

End Namespace

