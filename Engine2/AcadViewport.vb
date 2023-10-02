
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Drawing
Imports Autodesk.AutoCAD.EditorInput

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2018
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

'http://through-the-interface.typepad.com/through_the_interface/2015/11/creating-an-autocad-layout-with-custom-plot-and-viewport-settings-using-net.html
'http://docs.autodesk.com/MAP/2014/CHS/index.html?url=filesACD/GUID-61C22902-F63B-4204-86EC-FA37312D1B6E.htm,topicNumber=ACDd30e713927
Namespace Engine2

    Public Class AcadViewport

        ''' <summary>
        ''' Lista as viewports
        ''' </summary>
        ''' <param name="LayoutName">Nome do layout</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListViewports(LayoutName As String, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As List(Of Viewport)
            Dim Database As Database
            Dim Editor As Editor
            Dim LayoutManager As LayoutManager
            Dim DBDictionary As DBDictionary
            Dim Layout As Layout
            Dim Viewport As Viewport
            ListViewports = New List(Of Viewport)
            If IsNothing(Document) = True Then
                Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Else
                If ActiveDocument = True Then
                    If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
            End If
            Database = Document.Database
            Editor = Document.Editor
            If IsNothing(Transaction) = False Then
                LayoutManager = LayoutManager.Current
                DBDictionary = Transaction.GetObject(Database.LayoutDictionaryId, OpenMode.ForRead)
                Layout = Transaction.GetObject(CType(DBDictionary(LayoutName), ObjectId), OpenMode.ForRead)
                For Each ID As ObjectId In Layout.GetViewports
                    Viewport = Transaction.GetObject(ID, OpenMode.ForWrite)
                    ListViewports.Add(Viewport)
                Next
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        ListViewports = ListViewports(LayoutName, Document, Transaction, ActiveDocument)
                        Transaction.Commit()
                    End Using
                End Using
            End If
            Return ListViewports
        End Function

        ''' <summary>
        ''' Cria a nova viewport
        ''' </summary>
        ''' <param name="LayoutName">Nome do layout que receberá a viewport</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="Height">Altura</param>
        ''' <param name="Width">Comprimento</param>
        ''' <param name="ViewTableRecord">Informações do foco da viewport</param>
        ''' <param name="ScaleFactor">Escala</param>
        ''' <param name="Visible">Visibilidade do conteúdo</param>
        ''' <param name="On">Determina se deve estar ligada</param>
        ''' <param name="Locked">Determina se deve ser bloqueada</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <returns>Viewport</returns>
        ''' <remarks></remarks>
        Public Shared Function CreateViewport(LayoutName As String, CenterPoint As Point3d, Width As Double, Height As Double, ViewTableRecord As ViewTableRecord, Optional ScaleFactor As Double = 1, Optional Visible As Boolean = True, Optional [On] As Boolean = False, Optional Locked As Boolean = True, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As Viewport
            Dim Viewport As Viewport = Nothing
            Dim Database As Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Editor As Editor
            Dim DBDictionary As DBDictionary
            Dim Layout As Layout
            If IsNothing(Document) = True Then
                Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Else
                If ActiveDocument = True Then
                    If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
            End If
            Database = Document.Database
            Editor = Document.Editor
            If IsNothing(Transaction) = False Then
                DBDictionary = Transaction.GetObject(Database.LayoutDictionaryId, OpenMode.ForRead)
                Layout = Transaction.GetObject(CType(DBDictionary(LayoutName), ObjectId), OpenMode.ForRead)
                BlockTableRecord = Transaction.GetObject(Layout.BlockTableRecordId, OpenMode.ForWrite)
                Viewport = New Viewport()
                With Viewport
                    .StandardScale = StandardScaleType.Scale1To1
                    .CenterPoint = CenterPoint
                    .Width = Width
                    .Height = Height
                    BlockTableRecord.AppendEntity(Viewport)
                    Transaction.AddNewlyCreatedDBObject(Viewport, True)
                    .On = [On]
                    .Locked = Locked
                    .Visible = Visible
                    .ViewCenter = ViewTableRecord.CenterPoint
                    .ViewDirection = ViewTableRecord.ViewDirection
                    .ViewHeight = ViewTableRecord.Height
                    .CustomScale = ScaleFactor
                    .UpdateDisplay()
                End With
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Viewport = CreateViewport(LayoutName, CenterPoint, Width, Height, ViewTableRecord, ScaleFactor, Visible, [On], Locked, Document, Transaction, ActiveDocument)
                        Transaction.Commit()
                    End Using
                End Using
            End If
            Return Viewport
        End Function

        ''' <summary>
        ''' Redimensiona a viewport
        ''' </summary>
        ''' <param name="Viewport">Viewport</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="Height">Altura</param>
        ''' <param name="Width">Comprimento</param>
        ''' <param name="ViewTableRecord">Informações do foco da viewport</param>
        ''' <param name="ScaleFactor">Escala</param>
        ''' <param name="Visible">Visibilidade do conteúdo</param>
        ''' <param name="On">Determina se deve estar ligada</param>
        ''' <param name="Locked">Determina se deve ser bloqueada</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ResizeViewport(Viewport As Viewport, CenterPoint As Point3d, Width As Double, Height As Double, ViewTableRecord As ViewTableRecord, Optional ScaleFactor As Double = 1, Optional Visible As Boolean = True, Optional [On] As Boolean = False, Optional Locked As Boolean = True, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As Viewport
            Dim Database As Database
            Dim Editor As Editor
            If IsNothing(Document) = True Then
                Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Else
                If ActiveDocument = True Then
                    If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
            End If
            Database = Document.Database
            Editor = Document.Editor
            If IsNothing(Transaction) = False Then
                Viewport = Transaction.GetObject(Viewport.Id, OpenMode.ForWrite)
                With Viewport
                    .CenterPoint = CenterPoint
                    .Width = Width
                    .Height = Height
                    .On = [On]
                    .Locked = Locked
                    .Visible = Visible
                    .ViewCenter = ViewTableRecord.CenterPoint
                    .ViewDirection = ViewTableRecord.ViewDirection
                    .ViewHeight = ViewTableRecord.Height
                    .CustomScale = ScaleFactor
                    .UpdateDisplay()
                End With
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Viewport = ResizeViewport(Viewport, CenterPoint, Width, Height, ViewTableRecord, ScaleFactor, Visible, [On], Locked, Document, Transaction, ActiveDocument)
                        Transaction.Commit()
                    End Using
                End Using
            End If
            Return Viewport
        End Function

        ''' <summary>
        ''' Exclui todas as viewports de um layout
        ''' </summary>
        ''' <param name="LayoutName">Nome do layout</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <param name="Exception">Coleção de ID´s a serem preservados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseAllViewports(LayoutName As String, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True, Optional Exception As List(Of ObjectId) = Nothing) As Boolean
            Dim Database As Database
            Dim Editor As Editor
            Dim LayoutManager As LayoutManager
            Dim DBDictionary As DBDictionary
            Dim Layout As Layout
            Dim Viewport As Viewport
            If IsNothing(Document) = True Then
                Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Else
                If ActiveDocument = True Then
                    If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
            End If
            Database = Document.Database
            Editor = Document.Editor
            If IsNothing(Transaction) = False Then
                If IsNothing(Exception) = True Then
                    Exception = New List(Of ObjectId)
                End If
                LayoutManager = LayoutManager.Current
                DBDictionary = Transaction.GetObject(Database.LayoutDictionaryId, OpenMode.ForRead)
                Layout = Transaction.GetObject(CType(DBDictionary(LayoutName), ObjectId), OpenMode.ForRead)
                For Each ID As ObjectId In Layout.GetViewports
                    Viewport = Transaction.GetObject(ID, OpenMode.ForWrite)
                    If Exception.Contains(Viewport.ObjectId) = False Then
                        Viewport.Erase()
                    End If
                Next
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        EraseAllViewports(LayoutName, Document, Transaction, ActiveDocument)
                        Transaction.Commit()
                    End Using
                End Using
            End If
            Return True
        End Function

    End Class

End Namespace



