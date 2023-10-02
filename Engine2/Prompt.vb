Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Reflection
Imports System.Text
Imports Autodesk.AutoCAD


Namespace Engine2

    ''' <summary>
    ''' Classe de interface do usuário
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Prompt

        ''' <summary>
        ''' Solicita a seleção de uma entidade na tela
        ''' </summary>
        ''' <param name="Message"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEntity(ByVal Message As String) As Object
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim acadEditor As Editor = acadDocument.Editor
                Dim KeywordCollection As New KeywordCollection
                Dim SB As New StringBuilder
                Dim Split() As String
                Dim CurrentKeyword As String = ""
                With SB
                    .Append(vbLf)
                    .Append(Message)
                    If Message.Contains("[") = True Then
                        Split = Message.Split("[")(1).Split("]")(0).Split("/")
                        For ini As Integer = 0 To Split.GetLength(0) - 1
                            KeywordCollection.Add(Split(ini).Trim)
                        Next
                    End If
                    If Message.Contains("<") = True Then
                        CurrentKeyword = Message.Split("<")(1).Split(">")(0).Trim
                    End If
                End With
                Dim PromptEntityResult As PromptEntityResult
                Dim PromptEntityOptions As PromptEntityOptions = New PromptEntityOptions("")
                With PromptEntityOptions
                    .Message = SB.ToString
                    If CurrentKeyword.Trim <> "" Then
                        .AllowNone = True
                    Else
                        .AllowNone = False
                    End If
                    If IsNothing(KeywordCollection) = False Then
                        For ini As Integer = 0 To KeywordCollection.Count - 1
                            .Keywords.Add(KeywordCollection.Item(ini).GlobalName)
                        Next
                    End If
                End With
                PromptEntityResult = acadEditor.GetEntity(PromptEntityOptions)
                With PromptEntityResult
                    Select Case .Status
                        Case PromptStatus.OK
                            Return Engine2.ConvertObject.ObjectIDToDBObject(.ObjectId)
                        Case PromptStatus.Cancel
                            Return .Status.ToString
                        Case PromptStatus.Error
                            Throw New System.Exception("Ocorreu um erro inespecífico em UI.GetEntity")
                        Case PromptStatus.Keyword
                            Return .StringResult
                        Case PromptStatus.Modeless
                            Return .Status.ToString
                        Case PromptStatus.None
                            Return CurrentKeyword
                        Case PromptStatus.Other
                            Return .Status.ToString
                        Case Else
                            Throw New System.Exception("Ocorreu um erro em UI.GetEntity, valor de retorno inesperado para função")
                    End Select
                End With
            Catch ex As System.Exception
                Throw New System.Exception(ex.Message)
            End Try
        End Function

    End Class

End Namespace