Imports Autodesk.AutoCAD.EditorInput
Imports System.IO
Imports System.Windows.Media.Imaging
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Namespace Engine2

    ''' <summary>
    ''' Monitora a abertura do tooltip associado com monitoramento do cursor
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ToolTipMonitor : Implements IDisposable

        'Declarações (Controle)
        Private WithEvents _DocumentCollection As DocumentCollection = Nothing
        Private WithEvents _Editor As Editor = Nothing
        Private WithEvents _Document As Document = Nothing

        'Declarações
        Private _PointMonitorEventHandler As PointMonitorEventHandler
        Private _EventHandler As EventHandler = Nothing
        Private _ObjectIdCollection As ObjectIdCollection = Nothing
        Private _ToolTip As Autodesk.Internal.Windows.ToolTip = Nothing
        Private _RawPoint As Object

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me._RawPoint = Nothing
            Me._EventHandler = Nothing
            Me._PointMonitorEventHandler = Nothing
            Me._ObjectIdCollection = New ObjectIdCollection
            Me._ToolTip = Nothing
            Me._DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        End Sub

        ''' <summary>
        ''' Evento de abertura do tooltip
        ''' </summary>
        ''' <param name="Sender">ToolTip</param>
        ''' <param name="e">PointMonitorEventArgs</param>
        ''' <remarks></remarks>
        Public Event ToolTipMonitorOpened(Sender As Autodesk.Internal.Windows.ToolTip, e As ObjectIdCollection)

        ''' <summary>
        ''' Atualiza o controle
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Update()
            Me._ObjectIdCollection.Clear()
            Me._ToolTip = Nothing
            If IsNothing(Me._EventHandler) = False Then
                RemoveHandler Autodesk.Windows.ComponentManager.ToolTipOpened, Me._EventHandler
                Me._EventHandler = Nothing
            End If
            If IsNothing(Me._PointMonitorEventHandler) = False Then
                RemoveHandler Me._Editor.PointMonitor, Me._PointMonitorEventHandler
                Me._PointMonitorEventHandler = Nothing
            End If
            Me._Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If IsNothing(Me._Document) = False Then
                Me._Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Me._PointMonitorEventHandler = AddressOf PointMonitor
                AddHandler Me._Editor.PointMonitor, Me._PointMonitorEventHandler
                Me._EventHandler = AddressOf ToolTipOpened
                AddHandler Autodesk.Windows.ComponentManager.ToolTipOpened, Me._EventHandler
            Else
                Me._Editor = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Detecta quando o documento sofre por mudanças de foco e atualiza o controle
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _DocumentCollection_DocumentActivated(sender As Object, e As DocumentCollectionEventArgs) Handles _DocumentCollection.DocumentActivated
            Me.Update()
        End Sub

        ''' <summary>
        ''' Detecta quando o documento é destruído e atualiza o controle
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _DocumentCollection_DocumentDestroyed(sender As Object, e As DocumentDestroyedEventArgs) Handles _DocumentCollection.DocumentDestroyed
            Me.Update()
        End Sub

        ''' <summary>
        ''' Monitora o cursor
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub PointMonitor(sender As Object, e As PointMonitorEventArgs)
            If sender.Document.CommandInProgress = "" Then
                If e.Context.PointComputed = True Then
                    Me._RawPoint = e.Context.RawPoint
                Else
                    Me._RawPoint = Nothing
                End If
            Else
                Me._RawPoint = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Monitora a abertura do tooltip
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub ToolTipOpened(sender As Object, e As EventArgs)
            If IsNothing(sender) = False Then
                If IsNothing(Me._RawPoint) = False Then
                    Dim PromptNestedEntityOptions As PromptNestedEntityOptions
                    Dim PromptNestedEntityResult As PromptNestedEntityResult
                    PromptNestedEntityOptions = New PromptNestedEntityOptions("")
                    With PromptNestedEntityOptions
                        .NonInteractivePickPoint = CType(Me._RawPoint, Point3d)
                        .UseNonInteractivePickPoint = True
                    End With
                    PromptNestedEntityResult = Me._Editor.GetNestedEntity(PromptNestedEntityOptions)
                    If PromptNestedEntityResult.Status = PromptStatus.OK Then

                        RemoveHandler Me._Editor.PointMonitor, Me._PointMonitorEventHandler
                        Me._PointMonitorEventHandler = Nothing
                        RemoveHandler Autodesk.Windows.ComponentManager.ToolTipOpened, Me._EventHandler
                        Me._EventHandler = Nothing

                        Me._ToolTip = sender
                        If PromptNestedEntityResult.GetContainers.Length > 0 Then
                            Me._ObjectIdCollection = New ObjectIdCollection(PromptNestedEntityResult.GetContainers)
                        Else
                            Me._ObjectIdCollection.Clear()
                        End If
                        If Me._ObjectIdCollection.Contains(PromptNestedEntityResult.ObjectId) = False Then
                            Me._ObjectIdCollection.Insert(0, PromptNestedEntityResult.ObjectId)
                        End If
                        RaiseEvent ToolTipMonitorOpened(Me._ToolTip, Me._ObjectIdCollection)
                        Me._ToolTip = Nothing
                        Me._ObjectIdCollection.Clear()
                        Me._RawPoint = Nothing

                        Me._PointMonitorEventHandler = AddressOf PointMonitor
                        AddHandler Me._Editor.PointMonitor, Me._PointMonitorEventHandler
                        Me._EventHandler = AddressOf ToolTipOpened
                        AddHandler Autodesk.Windows.ComponentManager.ToolTipOpened, Me._EventHandler

                    Else
                        Me._ObjectIdCollection.Clear()
                        Me._RawPoint = Nothing
                    End If
                Else
                    Me._ObjectIdCollection.Clear()
                    Me._RawPoint = Nothing
                End If
            Else
                Me._ToolTip = Nothing
                Me._ObjectIdCollection.Clear()
                Me._RawPoint = Nothing
            End If
        End Sub

#Region "Shareds"

        ''' <summary>
        ''' Converte System.Drawing.Image em System.Windows.Controls.Image
        ''' </summary>
        ''' <param name="Image">System.Drawing.Image</param>
        ''' <param name="Format">Format</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DrawingImageToWindowsImage(Image As System.Drawing.Image, Format As System.Drawing.Imaging.ImageFormat) As System.Windows.Controls.Image
            DrawingImageToWindowsImage = New System.Windows.Controls.Image
            Dim BitMapImage As BitmapImage
            Dim MemoryStream As MemoryStream
            MemoryStream = New MemoryStream
            Image.Save(MemoryStream, Format)
            BitMapImage = New BitmapImage
            BitMapImage.BeginInit()
            MemoryStream.Seek(0, SeekOrigin.Begin)
            BitMapImage.StreamSource = MemoryStream
            BitMapImage.EndInit()
            DrawingImageToWindowsImage.Source = BitMapImage
            Return DrawingImageToWindowsImage
        End Function

#End Region

#Region "IDisposable Support"

        Private disposedValue As Boolean
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    If IsNothing(Me._EventHandler) = False Then
                        RemoveHandler Autodesk.Windows.ComponentManager.ToolTipOpened, Me._EventHandler
                        Me._EventHandler = Nothing
                    End If
                    If IsNothing(Me._PointMonitorEventHandler) = False Then
                        RemoveHandler Me._Editor.PointMonitor, Me._PointMonitorEventHandler
                        Me._PointMonitorEventHandler = Nothing
                    End If
                    Me._DocumentCollection = Nothing
                    Me._Document = Nothing
                    Me._Editor = Nothing
                    Me._ObjectIdCollection = Nothing
                    Me._ToolTip = Nothing
                    Me._RawPoint = Nothing
                    Me._ObjectIdCollection.Clear()
                    Me._ObjectIdCollection = Nothing
                End If
            End If
            Me.disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class

End Namespace


