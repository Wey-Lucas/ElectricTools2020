Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports ElectricTools2020.Engine2.Geometry

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Gerenciamento da UCS
    ''' </summary>
    ''' <remarks></remarks>
    Public Class UCS

        ''' <summary>
        ''' Valida CoordinateSystem3d
        ''' </summary>
        ''' <param name="CoordinateSystem3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsValidCoordinateSystem3d(CoordinateSystem3d As CoordinateSystem3d) As Boolean
            With CoordinateSystem3d
                If .Xaxis.X.Equals(0) = True And .Xaxis.Y.Equals(0) = True And .Xaxis.Z.Equals(0) = True And .Yaxis.X.Equals(0) = True And .Yaxis.Y.Equals(0) = True And .Yaxis.Z.Equals(0) = True And .Zaxis.X.Equals(0) = True And .Zaxis.Y.Equals(0) = True And .Xaxis.Z.Equals(0) = True Then
                    Return False
                End If
            End With
            Return True
        End Function

        ''' <summary>
        ''' Retorna CoordinateSystem3d corrente
        ''' </summary>
        ''' <returns>CoordinateSystem3d</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCoordinateSystem3d() As CoordinateSystem3d
            Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem.CoordinateSystem3d
        End Function

        ''' <summary>
        ''' Retorna CurrentUserCoordinateSystem
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentUserCoordinateSystem() As Matrix3d
            Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem
        End Function

        ''' <summary>
        ''' Rotaciona a UCS
        ''' </summary>
        ''' <param name="Angle">Ângulo (Radianos)</param>
        ''' <param name="Axis">Eixo</param>
        ''' <remarks></remarks>
        Public Shared Sub RotateUCS(Angle As Double, Axis As eAxis)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim ViewportTableRecord As ViewportTableRecord
            Dim CoordinateSystem3d As CoordinateSystem3d
            Using DocumentLock As DocumentLock = Editor.Document.LockDocument
                Using Transaction As Transaction = Editor.Document.Database.TransactionManager.StartTransaction
                    ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, OpenMode.ForWrite)
                    Select Case Axis
                        Case eAxis.X
                            Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Xaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
                        Case eAxis.Y
                            Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Yaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
                        Case eAxis.Z
                            Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
                    End Select
                    CoordinateSystem3d = Editor.CurrentUserCoordinateSystem.CoordinateSystem3d
                    ViewportTableRecord.SetUcs(CoordinateSystem3d.Origin, CoordinateSystem3d.Xaxis, CoordinateSystem3d.Yaxis)
                    Editor.UpdateTiledViewportsFromDatabase()
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Rotaciona a UCS
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Angle">Ângulo (Radianos)</param>
        ''' <param name="Axis">Eixo</param>
        ''' <remarks></remarks>
        Public Shared Sub RotateUCS(Transaction As Transaction, Angle As Double, Axis As eAxis)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim ViewportTableRecord As ViewportTableRecord
            Dim CoordinateSystem3d As CoordinateSystem3d
            ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, OpenMode.ForWrite)
            Select Case Axis
                Case eAxis.X
                    Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Xaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
                Case eAxis.Y
                    Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Yaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
                Case eAxis.Z
                    Editor.CurrentUserCoordinateSystem = Matrix3d.Rotation(Angle, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis, Editor.CurrentUserCoordinateSystem.CoordinateSystem3d.Origin)
            End Select
            CoordinateSystem3d = Editor.CurrentUserCoordinateSystem.CoordinateSystem3d
            ViewportTableRecord.SetUcs(CoordinateSystem3d.Origin, CoordinateSystem3d.Xaxis, CoordinateSystem3d.Yaxis)
            Editor.UpdateTiledViewportsFromDatabase()
        End Sub

        ''' <summary>
        ''' Ajusta a UCS
        ''' </summary>
        ''' <param name="CoordinateSystem3d">CoordinateSystem3d</param>
        ''' <remarks></remarks>
        Public Shared Sub SetUCS(CoordinateSystem3d As CoordinateSystem3d)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim ViewportTableRecord As ViewportTableRecord
            Using DocumentLock As DocumentLock = Editor.Document.LockDocument
                Using Transaction As Transaction = Editor.Document.Database.TransactionManager.StartTransaction
                    ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, OpenMode.ForWrite)
                    ViewportTableRecord.SetUcs(CoordinateSystem3d.Origin, CoordinateSystem3d.Xaxis, CoordinateSystem3d.Yaxis)
                    Editor.UpdateTiledViewportsFromDatabase()
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Ajusta a UCS
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="CoordinateSystem3d">CoordinateSystem3d</param>
        ''' <remarks></remarks>
        Public Shared Sub SetUCS(Transaction As Transaction, CoordinateSystem3d As CoordinateSystem3d)
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim ViewportTableRecord As ViewportTableRecord
            ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, OpenMode.ForWrite)
            ViewportTableRecord.SetUcs(CoordinateSystem3d.Origin, CoordinateSystem3d.Xaxis, CoordinateSystem3d.Yaxis)
            Editor.UpdateTiledViewportsFromDatabase()
        End Sub

        ''' <summary>
        ''' Restaura a UCS original (World)
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub UcsWorld()
            Dim ViewportTableRecord As ViewportTableRecord
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Using DocumentLock As DocumentLock = Editor.Document.LockDocument
                Using Transaction As Transaction = Editor.Document.Database.TransactionManager.StartTransaction
                    ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                    ViewportTableRecord.SetUcsToWorld()
                    Editor.UpdateTiledViewportsFromDatabase()
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Restaura a UCS original (World)
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <remarks></remarks>
        Public Shared Sub UcsWorld(Transaction As Transaction)
            Dim ViewportTableRecord As ViewportTableRecord
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            ViewportTableRecord = Transaction.GetObject(Editor.ActiveViewportId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            ViewportTableRecord.SetUcsToWorld()
            Editor.UpdateTiledViewportsFromDatabase()
        End Sub

        ''' <summary>
        ''' Converte ponto 3d da UCS para WCS
        ''' </summary>
        ''' <param name="UcsPoint3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point3dUcsToWcs(UcsPoint3d As Point3d) As Point3d
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Return UcsPoint3d.TransformBy(Editor.CurrentUserCoordinateSystem)
        End Function

        ''' <summary>
        ''' Converte ponto 3d da WCS para UCS
        ''' </summary>
        ''' <param name="UcsPoint3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point3dWcsToUcs(UcsPoint3d As Point3d) As Point3d
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Return UcsPoint3d.TransformBy(Editor.CurrentUserCoordinateSystem.Inverse)
        End Function

        ''' <summary>
        ''' Obtem o ângulo da UCS em relação a WCS
        ''' </summary>
        ''' <param name="Axis">Eixo</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetUcsAngle(Axis As eAxis, AngleFormat As eAngleFormat) As Double
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim CoordinateSystem3d As CoordinateSystem3d = Editor.CurrentUserCoordinateSystem.CoordinateSystem3d
            Select Case Axis
                Case eAxis.X
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.XAxis))
                Case eAxis.Y
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.YAxis))
                Case eAxis.Z
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.ZAxis))
            End Select
            Select Case AngleFormat
                Case eAngleFormat.Degrees
                    GetUcsAngle = Engine2.Geometry.RadianToDegree(GetUcsAngle, True)
            End Select
            Return GetUcsAngle
        End Function

        ''' <summary>
        ''' Obtem o ângulo da UCS em relação a WCS
        ''' </summary>
        ''' <param name="CoordinateSystem3d">CoordinateSystem3d</param>
        ''' <param name="Axis">Eixo</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetUcsAngle(CoordinateSystem3d As CoordinateSystem3d, Axis As eAxis, AngleFormat As eAngleFormat) As Double
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Select Case Axis
                Case eAxis.X
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.XAxis))
                Case eAxis.Y
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.YAxis))
                Case eAxis.Z
                    GetUcsAngle = CoordinateSystem3d.Xaxis.AngleOnPlane(New Plane(CoordinateSystem3d.Origin, Vector3d.ZAxis))
            End Select
            Select Case AngleFormat
                Case eAngleFormat.Degrees
                    GetUcsAngle = Engine2.Geometry.RadianToDegree(GetUcsAngle, True)
            End Select
            Return GetUcsAngle
        End Function

#Region "ESTUDOS"

        ' ''' <summary>
        ' ''' Restaura WCS
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Shared Sub RestoreWCS()
        '    Dim Document As Document = Application.DocumentManager.MdiActiveDocument
        '    Dim Editor As Editor = Document.Editor
        '    Editor.CurrentUserCoordinateSystem = Matrix3d.Identity
        '    Editor.Regen()
        'End Sub

        'Public Shared Sub SetUCS(DegreeAngle As Double)
        '    Dim db As Database = HostApplicationServices.WorkingDatabase
        '    Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        '    Dim options As New PromptEntityOptions(vbLf & "Select an Line Entity to set the UCS too : ")
        '    options.SetRejectMessage(vbLf & "Select Line")
        '    options.AddAllowedClass(GetType(Line), False)
        '    Dim res As PromptEntityResult = ed.GetEntity(options)
        '    If res.Status <> PromptStatus.OK Then
        '        Return
        '    End If
        '    Using Tx As Transaction = db.TransactionManager.StartTransaction()
        '        Dim ent As Line = TryCast(Tx.GetObject(res.ObjectId, OpenMode.ForRead), Line)
        '        Dim ecs As New Matrix3d()
        '        Dim xVec As Vector3d = (ent.EndPoint - ent.StartPoint)
        '        xVec = xVec.GetNormal()
        '        Dim coordSystem As New CoordinateSystem3d(ent.StartPoint, xVec, xVec.CrossProduct(ent.Normal).Negate())
        '        Dim cosy As CoordinateSystem3d = ed.CurrentUserCoordinateSystem.CoordinateSystem3d
        '        ecs = Matrix3d.AlignCoordinateSystem(cosy.Origin, cosy.Xaxis, cosy.Yaxis, cosy.Zaxis, coordSystem.Origin, coordSystem.Xaxis, coordSystem.Yaxis, coordSystem.Zaxis)
        '        ed.CurrentUserCoordinateSystem = ecs
        '        Tx.Commit()
        '    End Using
        '    ed.UpdateTiledViewportsInDatabase()

        'End Sub

        'Private Sub CreateUCSAlignedToEntity(surfaceId As ObjectId, editor As Editor, dwg As Database)
        '    ' We're going to make an inquiry into the drawing DB so start a transaction
        '    Dim trans As Transaction = editor.Document.Database.TransactionManager.StartTransaction()
        '    Try
        '        ' Change the Current UCS to that of the 'surface' we want to draw the panels on.
        '        Dim entity As Entity = DirectCast(trans.GetObject(surfaceId, OpenMode.ForRead), Entity)

        '        ' Gather the necessary data that defines the entity's unique coordinate system
        '        Dim entityOrigin As Point3d = entity.GeometricExtents.MinPoint
        '        Dim entityXaxis As Vector3d = entity.Ecs.CoordinateSystem3d.Xaxis
        '        Dim entityYaxis As Vector3d = entity.Ecs.CoordinateSystem3d.Yaxis

        '        ' Get the Active Viewport
        '        Dim viewportTableRecord As ViewportTableRecord = DirectCast(trans.GetObject(editor.ActiveViewportId, OpenMode.ForWrite), ViewportTableRecord)
        '        viewportTableRecord.IconAtOrigin = True
        '        viewportTableRecord.IconEnabled = True

        '        ' Debug: Neither of these two statements properly rotates the UCS in 3D space!
        '        'editor.CurrentUserCoordinateSystem = entity.Ecs;
        '        editor.CurrentUserCoordinateSystem = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, entityOrigin, entityXaxis, entityYaxis, entityXaxis.CrossProduct(entityYaxis))

        '        viewportTableRecord.SetUcs(entityOrigin, entityXaxis, entityYaxis)
        '        editor.UpdateTiledViewportsFromDatabase()

        '        trans.Commit()

        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        editor.WriteMessage(vbLf & "Error: " + ex.Message)
        '    Finally

        '        trans.Dispose()
        '    End Try
        'End Sub

        'Public Shared Sub UcsWorld()
        '    Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        '    Using View As ViewTableRecord = Editor.GetCurrentView()
        '        Editor.CurrentUserCoordinateSystem = Matrix3d.PlaneToWorld(View.ViewDirection) * Matrix3d.Rotation(-View.ViewTwist, Vector3d.ZAxis, Point3d.Origin)
        '    End Using
        'End Sub

#End Region

    End Class

End Namespace


