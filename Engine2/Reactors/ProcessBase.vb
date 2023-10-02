Imports Autodesk.AutoCAD.DatabaseServices

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2016
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe base para criação de '...EventArgs' dos eventos em reators
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class ProcessBase

        ''' <summary>
        ''' Classe proprietária
        ''' </summary>
        ''' <remarks></remarks>
        Private _Reactors As Reactors

        ''' <summary>
        ''' Estado original da entidade
        ''' </summary>
        ''' <remarks></remarks>
        Private _OriginalState As Entity

        ''' <summary>
        ''' Id original da entidade
        ''' </summary>
        ''' <remarks></remarks>
        Private _ObjectId As ObjectId

        ''' <summary>
        ''' Estado inicial da entidade
        ''' </summary>
        ''' <remarks></remarks>
        Private _OldState As Entity

        ''' <summary>
        ''' Id original da entidade
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
        ''' Estado corrente da entidade
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Entity As Entity
            Get
                Return Me._ObjectId.ToEntity
            End Get
        End Property

        ''' <summary>
        ''' Estado original da entidade
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property OriginalState As Entity
            Get
                Return Me._OriginalState
            End Get
        End Property

        ''' <summary>
        ''' Estado inicial da entidade (Clone)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property OldState As Entity
            Get
                Return Me._OldState
            End Get
        End Property

        ''' <summary>
        ''' Reactor
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Reactor As Reactor
            Get
                Return Me._Reactors.Find(Function(X As Reactor) X.Entity.ObjectId = Me._ObjectId)
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Reactors">Classe proprietária</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef Reactors As Reactors, ObjectId As ObjectId)
            Me._Reactors = Reactors
            Me._ObjectId = ObjectId
            Me._OriginalState = Me._ObjectId.ToEntity
            Me._OldState = Me._OriginalState.Clone
        End Sub

    End Class

End Namespace