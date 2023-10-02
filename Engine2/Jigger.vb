Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface

Namespace Engine2

    ''' <summary>
    ''' Fantasma para Criar círculo a partir do raio
    ''' </summary>
    ''' <remarks></remarks>
    Public Class JigCircle : Inherits DrawJig

        ''' <summary>
        ''' Armazena o ponto de referência
        ''' </summary>
        ''' <remarks></remarks>
        Private _Center As Point3d

        ''' <summary>
        ''' Armazena o ponto de destino
        ''' </summary>
        ''' <remarks></remarks>
        Private _RadiusPoint As Point3d

        ''' <summary>
        ''' Raio
        ''' </summary>
        Private _Radius As Double

        ''' <summary>
        ''' UCS
        ''' </summary>
        Private _UCS As Matrix3d

        ''' <summary>
        ''' Mensagem de seleção para ponto destino
        ''' </summary>
        ''' <remarks></remarks>
        Private _PromptRadiusPointMessage As String

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Center">Centro do cículo</param>
        ''' <param name="PromptRadiusPointMessage">Mensagem para o ponto de definição do raio</param>
        Public Sub New(Center As Point3d, PromptRadiusPointMessage As String)
            Me._UCS = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem
            Me._Center = Center.TransformBy(Me._UCS)
            Me._RadiusPoint = New Point3d(0, 0, 0)
            Me._PromptRadiusPointMessage = PromptRadiusPointMessage
        End Sub

        ''' <summary>
        ''' Radius
        ''' </summary>
        ''' <returns></returns>
        Public Property Radius As Double
            Get
                Return Me._Radius
            End Get
            Set(ByVal value As Double)
                Me._Radius = value
            End Set
        End Property

        ''' <summary>
        ''' Calcula a movimentação
        ''' </summary>
        ''' <param name="draw"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
            Dim geo As WorldGeometry = draw.Geometry
            If geo IsNot Nothing Then
                geo.PushModelTransform(Me._UCS)
                geo.Circle(Me._Center, Me._Radius, Vector3d.ZAxis)
                geo.PopModelTransform()
            End If
            Return True
        End Function

        ''' <summary>
        ''' Exibe a mensagem de seleção
        ''' </summary>
        ''' <param name="prompts"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function Sampler(ByVal prompts As JigPrompts) As SamplerStatus
            Dim JigPromptPointOptions As JigPromptPointOptions = New JigPromptPointOptions(vbLf & Me._PromptRadiusPointMessage)
            JigPromptPointOptions.BasePoint = Me._Center
            JigPromptPointOptions.UseBasePoint = True
            Dim PromptPointResult As PromptPointResult = prompts.AcquirePoint(JigPromptPointOptions)
            Select Case PromptPointResult.Status
                Case PromptStatus.Cancel, PromptStatus.Error
                    Return SamplerStatus.Cancel
                Case Else
                    If Me._Center.DistanceTo(PromptPointResult.Value) > Tolerance.Global.EqualPoint Then
                        Me._RadiusPoint = PromptPointResult.Value.TransformBy(Me._UCS.Inverse())
                        Me._Radius = Me._RadiusPoint.DistanceTo(Me._Center)
                        Return SamplerStatus.OK
                    Else
                        Return SamplerStatus.NoChange
                    End If
            End Select
        End Function

    End Class

    ''' <summary>
    ''' Fantasma para Mover ou copiar entidades
    ''' </summary>
    ''' <remarks></remarks>
    Public Class JigCopyOrMoveSelection : Inherits DrawJig

        ''' <summary>
        ''' Armazena o ponto de referência
        ''' </summary>
        ''' <remarks></remarks>
        Private _PointReference As Point3d

        ''' <summary>
        ''' Armazena o ponto de destino
        ''' </summary>
        ''' <remarks></remarks>
        Private _PointDisplacement As Point3d

        ''' <summary>
        ''' Armazena a coleção de entidades
        ''' </summary>
        ''' <remarks></remarks>
        Private _EntityCollection As List(Of Entity)

        ''' <summary>
        ''' Mensagem de seleção para ponto destino
        ''' </summary>
        ''' <remarks></remarks>
        Private _PromptDisplacementMessage As String

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PromptDisplacementMessage">Mensagem de seleção para ponto destino</param>
        ''' <remarks></remarks>
        Public Sub New(PointReference As Point3d, PromptDisplacementMessage As String)
            Dim UCS As Matrix3d = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem
            Me._EntityCollection = New List(Of Entity)
            Me._PointReference = PointReference.TransformBy(UCS)
            Me._PromptDisplacementMessage = PromptDisplacementMessage
        End Sub

        ''' <summary>
        ''' Ponto de referência
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PointReference() As Point3d
            Get
                Return Me._PointReference
            End Get
        End Property

        ''' <summary>
        ''' Ponto de destino
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PointDisplacement() As Point3d
            Get
                Return Me._PointDisplacement
            End Get
        End Property

        ''' <summary>
        ''' Adiciona a coleção de ID dos itens a serem movidos ou copiados
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s</param>
        ''' <remarks></remarks>
        Public Sub Add(ObjectIdCollection As ObjectIdCollection)
            Dim Entity As Entity
            For Each ObjectId As ObjectId In ObjectIdCollection
                Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
                Me._EntityCollection.Add(Entity)
            Next
        End Sub

        ''' <summary>
        ''' Adiciona a entidade 
        ''' </summary>
        ''' <param name="ObjectID">ID da entidade a ser movida</param>
        ''' <remarks></remarks>
        Public Sub Add(ObjectID As ObjectId)
            Dim Entity As Entity
            Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectID)
            Me._EntityCollection.Add(Entity)
        End Sub

        ''' <summary>
        ''' Move o conjunto para o ponto destino e retorna a coleção modificada
        ''' </summary>
        ''' <remarks></remarks>
        Public Function TransformMoveEntities() As List(Of Entity)
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(_PointReference.GetVectorTo(_PointDisplacement))
            For Each Entity As Entity In Me._EntityCollection
                Entity.TransformBy(Matrix3d)
            Next
            Return Me._EntityCollection
        End Function

        ''' <summary>
        ''' Copia o conjunto para o ponto destino e retorna a coleção modificada
        ''' </summary>
        ''' <remarks></remarks>
        Public Function TransformCopyEntities() As List(Of Entity)
            Dim Clone As Entity
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(_PointReference.GetVectorTo(_PointDisplacement))
            Dim EntityCopyCollection As New List(Of Entity)
            For Each Entity As Entity In Me._EntityCollection
                Clone = Entity.Clone
                Clone.TransformBy(Matrix3d)
                EntityCopyCollection.Add(Clone)
            Next
            Return EntityCopyCollection
        End Function

        ''' <summary>
        ''' Calcula a movimentação
        ''' </summary>
        ''' <param name="draw"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function WorldDraw(draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(_PointReference.GetVectorTo(_PointDisplacement))
            Dim WorldGeometry As WorldGeometry = draw.Geometry
            If IsNothing(WorldGeometry) = False Then
                WorldGeometry.PushModelTransform(Matrix3d)
                For Each Entity As Entity In Me._EntityCollection
                    WorldGeometry.Draw(Entity)
                Next
                WorldGeometry.PopModelTransform()
            End If
            Return True
        End Function

        ''' <summary>
        ''' Exibe a mensagem de seleção
        ''' </summary>
        ''' <param name="prompts"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function Sampler(prompts As JigPrompts) As SamplerStatus
            Dim JigPromptPointOptions As New JigPromptPointOptions(vbLf & _PromptDisplacementMessage)
            JigPromptPointOptions.UseBasePoint = False
            Dim PromptPointResult As PromptPointResult = prompts.AcquirePoint(JigPromptPointOptions)
            Select Case PromptPointResult.Status
                Case PromptStatus.Cancel, PromptStatus.Error
                    Return SamplerStatus.Cancel
                Case Else
                    If Me._PointDisplacement.DistanceTo(PromptPointResult.Value) > Tolerance.Global.EqualPoint Then
                        Me._PointDisplacement = PromptPointResult.Value
                        Return SamplerStatus.OK
                    Else
                        Return SamplerStatus.NoChange
                    End If
            End Select
        End Function

    End Class


End Namespace
