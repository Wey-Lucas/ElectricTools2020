Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Namespace Engine2

    ''' <summary>
    ''' Inibe a reedição de pontos pelo Grip
    ''' Utilizado por Reactor
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MyGripOverrule : Inherits GripOverrule

        ''' <summary>
        ''' Instancia o dicionário do sistema
        ''' </summary>
        ''' <remarks></remarks>
        Private Shared _Dictionary As New Dictionary(Of String, Point3dCollection)()

        ''' <summary>
        ''' Determina se a classe esta ativa
        ''' </summary>
        ''' <remarks></remarks>
        Private _Running As Boolean

        ''' <summary>
        ''' Obtem a chave do dicionário
        ''' </summary>
        ''' <param name="e"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetKey(e As Entity) As String
            Return e.GetType.Name & ":" & e.GeometricExtents.ToString()
        End Function

        ''' <summary>
        ''' Armazenagem de pontos do dicionário
        ''' </summary>
        ''' <param name="e"></param>
        ''' <param name="grips"></param>
        ''' <remarks></remarks>
        Private Sub StoreGripInfo(e As Entity, grips As Point3dCollection)
            Dim key As String = GetKey(e)
            If _Dictionary.ContainsKey(key) Then
                Using grps As Point3dCollection = _Dictionary(key)
                    grps.Clear()
                End Using
                _Dictionary.Remove(key)
            End If
            Dim pts As Point3d() = New Point3d(grips.Count - 1) {}
            grips.CopyTo(pts, 0)
            Dim gps As New Point3dCollection(pts)
            _Dictionary.Add(key, gps)
        End Sub

        ''' <summary>
        ''' Informações do grip
        ''' </summary>
        ''' <param name="e"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function RetrieveGripInfo(e As Entity) As Point3dCollection
            Dim grips As Point3dCollection = Nothing
            Dim key As String = GetKey(e)
            If _Dictionary.ContainsKey(key) Then
                grips = _Dictionary(key)
            End If
            Return grips
        End Function

        ''' <summary>
        ''' Obtem os pontos grip
        ''' </summary>
        ''' <param name="e">Entidade</param>
        ''' <param name="grips">Coleção de pontos</param>
        ''' <param name="snaps">Snaps</param>
        ''' <param name="geomIds">ID´s da geometria</param>
        ''' <remarks></remarks>
        Public Overrides Sub GetGripPoints(e As Entity, grips As Point3dCollection, snaps As IntegerCollection, geomIds As IntegerCollection)
            MyBase.GetGripPoints(e, grips, snaps, geomIds)
            StoreGripInfo(e, grips)
        End Sub

        ''' <summary>
        ''' Obtem a movimentação dos grips
        ''' </summary>
        ''' <param name="e">Entidade</param>
        ''' <param name="indices">Índices</param>
        ''' <param name="offset">offset</param>
        ''' <remarks></remarks>
        Public Overrides Sub MoveGripPointsAt(e As Entity, indices As IntegerCollection, offset As Vector3d)
            Dim Grips As Point3dCollection = RetrieveGripInfo(e)
            If IsNothing(Grips) = False Then
                MyBase.MoveGripPointsAt(e, indices, offset)
            End If
        End Sub

        ''' <summary>
        ''' Liga o monitoramento de grip
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub GripOn()
            If Me._Running = False Then
                ObjectOverrule.AddOverrule(RXClass.GetClass(GetType(Entity)), New MyGripOverrule, True)
            End If
        End Sub

        ''' <summary>
        ''' Desliga o monitoramento de grip
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub GripOff()
            If Me._Running = True Then
                ObjectOverrule.RemoveOverrule(RXClass.GetClass(GetType(Entity)), New MyGripOverrule)
            End If
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me._Running = False
        End Sub

    End Class

End Namespace


