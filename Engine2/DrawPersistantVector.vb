Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.GraphicsInterface

Namespace Engine2

    ''' <summary>
    ''' Gerencia vetores persistentes no desenho
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DrawPersistantVector

        ''' <summary>
        ''' Declarações gerais
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents acadDocument As Document

        ''' <summary>
        ''' Armazena os graficos persistentes
        ''' </summary>
        ''' <remarks></remarks>
        Private _Markers As List(Of MyTransient)

        ''' <summary>
        ''' Limpa todos os vetores do desenho
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ClearAll()
            Dim TransientManager As TransientManager
            If IsNothing(_Markers) = False Then
                TransientManager = TransientManager.CurrentTransientManager
                For Each Marker As MyTransient In _Markers
                    TransientManager.EraseTransient(Marker, New IntegerCollection())
                    Marker.Dispose()
                Next
                _Markers.Clear()
            End If
        End Sub

        ''' <summary>
        ''' Limpa vetores específicos do desenho
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Clear(ErasedCollection As DBObjectCollection)
            Dim TransientManager As TransientManager
            If IsNothing(_Markers) = False Then
                TransientManager = TransientManager.CurrentTransientManager
                For Each Marker As MyTransient In ErasedCollection
                    If _Markers.Contains(Marker) = True Then
                        _Markers.Remove(Marker)
                        TransientManager.EraseTransient(Marker, New IntegerCollection())
                        Marker.Dispose()
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Desenha os vetores
        ''' </summary>
        ''' <param name="Drawable">Entidade a ser desenhada</param>
        ''' <param name="Highlight">Modo de destaque</param>
        ''' <param name="ClearOldVectors">Determina se vetores anteriores serão apagados</param>
        ''' <remarks></remarks>
        Public Sub DrawVector(Drawable As Entity, Highlight As Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode, Optional ClearOldVectors As Boolean = False)
            Dim TransientManager As TransientManager
            Dim IntegerCollection As New IntegerCollection
            Dim MyTransient As MyTransient
            If ClearOldVectors = True Then
                Me.ClearAll()
            End If
            MyTransient = New MyTransient(Drawable)
            If _Markers.Contains(MyTransient) = False Then
                _Markers.Add(MyTransient)
            End If
            Transient.CapturedDrawable = MyTransient
            TransientManager = TransientManager.CurrentTransientManager
            TransientManager.AddTransient(MyTransient, Highlight, 128, IntegerCollection)
        End Sub

        ''' <summary>
        ''' Desenha os vetores
        ''' </summary>
        ''' <param name="DrawableCollection">Coleção de entidades a serem desenhadas</param>
        ''' <param name="Highlight">Modo de destaque</param>
        ''' <param name="ClearOldVectors">Determina se vetores anteriores serão apagados</param>
        ''' <remarks></remarks>
        Public Sub DrawVector(DrawableCollection As List(Of Entity), Highlight As Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode, Optional ClearOldVectors As Boolean = False)
            Dim TransientManager As TransientManager
            Dim IntegerCollection As New IntegerCollection
            Dim MyTransient As MyTransient
            If ClearOldVectors = True Then
                Me.ClearAll()
            End If
            For Each Drawable As Drawable In DrawableCollection
                MyTransient = New MyTransient(Drawable)
                Transient.CapturedDrawable = MyTransient
                TransientManager = TransientManager.CurrentTransientManager
                TransientManager.AddTransient(MyTransient, Highlight, 128, IntegerCollection)
                If _Markers.Contains(MyTransient) = False Then
                    _Markers.Add(MyTransient)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Limpa os gráficos a partir do comando Redraw
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub acadDocument_CommandEnded(sender As Object, e As CommandEventArgs) Handles acadDocument.CommandEnded
            If e.GlobalCommandName.Equals("REDRAW") = True Then
                Me.ClearAll()
            End If
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            acadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Me._Markers = New List(Of MyTransient)
        End Sub

    End Class

    ''' <summary>
    ''' Complemento de DrawPersistantVector
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class MyTransient : Inherits Transient

        'Declarações
        Private _Entity As Entity = Nothing

        ''' <summary>
        ''' Seta o trato do attributo
        ''' </summary>
        ''' <param name="traits"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function SubSetAttributes(traits As DrawableTraits) As Integer
            Return CInt(DrawableAttributes.None)
        End Function

        ''' <summary>
        ''' Detecta o SubDraw
        ''' </summary>
        ''' <param name="vd"></param>
        ''' <remarks></remarks>
        Protected Overrides Sub SubViewportDraw(vd As ViewportDraw)
            Me._Entity.ViewportDraw(vd)
        End Sub

        ''' <summary>
        ''' Determina a cor da mensagem
        ''' </summary>
        ''' <param name="wd"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Function SubWorldDraw(wd As WorldDraw) As Boolean
            Me._Entity.WorldDraw(wd)
            wd.SubEntityTraits.Color = Me._Entity.Color.ColorIndex
            Return True
        End Function

        ''' <summary>
        ''' Detecta a entrada do cursor
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Protected Overrides Sub OnDeviceInput(e As DeviceInputEventArgs)
            MyBase.OnDeviceInput(e)
        End Sub

        ''' <summary>
        ''' Formador de mensagem
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ForceMessage()
        End Sub

        ''' <summary>
        ''' Detecta a posição do cursor
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Protected Overrides Sub OnPointInput(e As PointInputEventArgs)
            MyBase.OnPointInput(e)
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Entity"></param>
        ''' <remarks></remarks>
        Public Sub New(Entity As Entity)
            Me._Entity = Entity
        End Sub

    End Class

End Namespace

