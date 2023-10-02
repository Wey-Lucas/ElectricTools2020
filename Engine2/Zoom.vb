'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports System.Text

Namespace Engine2

    ''' <summary>
    ''' Zoom em arquivos DWG
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Zoom

        ''' <summary>
        ''' ZoomExtents
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomExtents(Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            Dim MaxPoint As Point3d
            Dim MinPoint As Point3d
            Dim MaxPoint2D As Point2d
            Dim MinPoint2D As Point2d
            Dim ViewTableRecord As ViewTableRecord
            If IsNothing(Transaction) = False Then
                Database.UpdateExt(True)
                MaxPoint = Database.Extmax
                MinPoint = Database.Extmin
                MaxPoint2D = New Point2d(MaxPoint.X, MaxPoint.Y)
                MinPoint2D = New Point2d(MinPoint.X, MinPoint.Y)
                ViewTableRecord = New ViewTableRecord
                ViewTableRecord.CenterPoint = MinPoint2D + (MaxPoint2D - MinPoint2D) * 0.5
                ViewTableRecord.Height = MaxPoint2D.Y - MinPoint2D.Y
                ViewTableRecord.Width = MaxPoint2D.X - MinPoint2D.X
                Editor.SetCurrentView(ViewTableRecord)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomExtents(Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Zoom object
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Transaction">Transação</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomObject(ByVal Entity As Entity, Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            Dim Extents3d As Extents3d
            If IsNothing(Transaction) = False Then
                Extents3d = Entity.GeometricExtents
                ZoomToWindow(New Point2d(Extents3d.MinPoint.X, Extents3d.MinPoint.Y), New Point2d(Extents3d.MaxPoint.X, Extents3d.MaxPoint.Y))
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomObject(Entity, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Zoom object
        ''' </summary>
        ''' <param name="ObjectIds">Coleção de ID´s</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomObjects(ObjectIds As List(Of ObjectId), Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            If IsNothing(Transaction) = False Then
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(ObjectIds.ToArray)
                Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomObjects(ObjectIds, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Zoom object
        ''' </summary>
        ''' <param name="ObjectIds">Coleção de ID´s</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomObjects(ObjectIds As ObjectIdCollection, Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            If IsNothing(Transaction) = False Then
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(ObjectIds.ObjectIdCollectionToListOfObjectId.ToArray)
                Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomObjects(ObjectIds, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' ZoomToWindow
        ''' </summary>
        ''' <param name="MinPoint">Ponto 1</param>
        ''' <param name="MaxPoint">Ponto 2</param>
        ''' <param name="Transaction">Transação</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomToWindow(ByVal MinPoint As Point2d, ByVal MaxPoint As Point2d, Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            Dim ViewTableRecord As ViewTableRecord
            If IsNothing(Transaction) = False Then
                ViewTableRecord = Editor.GetCurrentView()
                ViewTableRecord.Width = MaxPoint.X - MinPoint.X
                ViewTableRecord.Height = MaxPoint.Y - MinPoint.Y
                ViewTableRecord.CenterPoint = New Point2d(MinPoint.X + (ViewTableRecord.Width / 2), MinPoint.Y + (ViewTableRecord.Height / 2))
                Editor.SetCurrentView(ViewTableRecord)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomToWindow(MinPoint, MaxPoint, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Zoom window
        ''' </summary>
        ''' <param name="CurrentViewState">CurrentViewState</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomToWindow(CurrentViewState As CurrentViewState, Optional Transaction As Transaction = Nothing)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            Dim ViewTableRecord As ViewTableRecord
            If IsNothing(Transaction) = False Then
                ViewTableRecord = Editor.GetCurrentView()
                ViewTableRecord.Width = CurrentViewState.Width
                ViewTableRecord.Height = CurrentViewState.Height
                ViewTableRecord.CenterPoint = CurrentViewState.CenterPoint
                Editor.SetCurrentView(ViewTableRecord)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ZoomToWindow(CurrentViewState, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Obtem os pontos da view atual para aplicação no comando ZoomWindow
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>CurrentViewState</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrentViewState(Optional Transaction As Transaction = Nothing) As CurrentViewState
            GetCurrentViewState = Nothing
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Database As Database = Editor.Document.Database
            Dim Document As Document = Editor.Document
            Dim ViewTableRecord As ViewTableRecord
            If IsNothing(Transaction) = False Then
                ViewTableRecord = Editor.GetCurrentView()
                GetCurrentViewState = New CurrentViewState(ViewTableRecord.Width, ViewTableRecord.Height, ViewTableRecord.CenterPoint)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            GetCurrentViewState = GetCurrentViewState(Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using

            End If
            Return GetCurrentViewState
        End Function

    End Class

    ''' <summary>
    ''' Armazena as informações sobre a visualização atual da área de desenho do AutoCAD
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class CurrentViewState

        ''' <summary>
        ''' Comprimento
        ''' </summary>
        ''' <remarks></remarks>
        Private _Width As Double

        ''' <summary>
        ''' Altura
        ''' </summary>
        ''' <remarks></remarks>
        Private _Height As Double

        ''' <summary>
        ''' Ponto central
        ''' </summary>
        ''' <remarks></remarks>
        Private _CenterPoint As Point2d

        ''' <summary>
        ''' Comprimento
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Width As Double
            Get
                Return Me._Width
            End Get
        End Property

        ''' <summary>
        ''' Altura
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Height As Double
            Get
                Return Me._Height
            End Get
        End Property

        ''' <summary>
        ''' Ponto central
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CenterPoint As Point2d
            Get
                Return Me._CenterPoint
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Width">Comprimento</param>
        ''' <param name="Height">Altura</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <remarks></remarks>
        Public Sub New(Width As Double, Height As Double, CenterPoint As Point2d)
            Me._Width = Width
            Me._Height = Height
            Me._CenterPoint = CenterPoint
        End Sub

    End Class


End Namespace

