Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Text

Public Class CallCommand

    ''' <summary>
    ''' Carrega os comandos
    ''' </summary>
    ''' <param name="Command"></param>
    ''' <remarks></remarks>
    Public Shared Sub CallCommand(Command As String)
        Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim SB As New StringBuilder
        With SB
            .Clear()
            .Append(Command)
            .Append(" ")
        End With
        Document.SendStringToExecute(SB.ToString, True, False, False)
    End Sub

End Class
