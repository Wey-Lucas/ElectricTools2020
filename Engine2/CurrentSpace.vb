'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.GraphicsInterface

Namespace Engine2

    ''' <summary>
    ''' Classe para obter o espaço corrente do desenho
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class CurrentSpace

        ''' <summary>
        ''' Retorna o espaço corrente do desenho (Model ou Paper space)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentSpace() As String
            If LayoutManager.Current.CurrentLayout = "Model" Then
                Return BlockTableRecord.ModelSpace
            Else
                Return BlockTableRecord.PaperSpace
            End If
        End Function

    End Class

End Namespace