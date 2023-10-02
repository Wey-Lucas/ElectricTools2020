'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices

Namespace Engine2

    ''' <summary>
    ''' Classe para fechamento de polyline
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClosePolyline

        ''' <summary>
        ''' Fecha uma polyline
        ''' </summary>
        ''' <param name="DBObject">DBObject do tipo Polyline</param>
        ''' <remarks></remarks>
        Public Shared Function Close(DBObject As DBObject) As Boolean
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Select Case DBObject.GetType.Name.ToUpper
                Case "POLYLINE"
                    Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        Using Transaction As Transaction = acadDocument.Database.TransactionManager.StartTransaction()
                            Dim Polyline As Polyline = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForRead)
                            If Polyline.Closed = False Then
                                Polyline.UpgradeOpen()
                                Polyline.Closed = True
                                Transaction.Commit()
                            End If
                            Transaction.Dispose()
                        End Using
                    End Using
                    Return True
                Case "POLYLINE3D"
                    Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        Using Transaction As Transaction = acadDocument.Database.TransactionManager.StartTransaction()
                            Dim Polyline3D As Polyline3d = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForRead)
                            If Polyline3D.Closed = False Then
                                Polyline3D.UpgradeOpen()
                                Polyline3D.Closed = True
                                Transaction.Commit()
                            End If
                            Transaction.Dispose()
                        End Using
                    End Using
                Case Else
                    Throw New System.Exception("A entidade informada não é válida")
            End Select
        End Function

        ''' <summary>
        ''' Verifica se a Polyline é  fechada
        ''' </summary>
        ''' <param name="DBObject">DBObject do tipo Polyline</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsClosed(DBObject As DBObject) As Boolean
            Select Case DBObject.GetType.Name.ToUpper
                Case "POLYLINE"
                    Dim Polyline As Polyline = Engine2.ConvertObject.DBObjectToEntity(DBObject)
                    Return Polyline.Closed
                Case "POLYLINE3D"
                    Dim Polyline3D As Polyline3d = Engine2.ConvertObject.DBObjectToEntity(DBObject)
                    Return Polyline3D.Closed
                Case Else
                    Throw New System.Exception("A entidade informada não é válida")
            End Select
        End Function

    End Class

End Namespace


