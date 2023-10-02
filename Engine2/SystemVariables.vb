Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe para consulta de variáveis
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class SystemVariables

        ''' <summary>
        ''' Retorna a coleção de variáveis do sistema
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSysVars() As List(Of SysVar)
            GetSysVars = New List(Of SysVar)
            For Each Variable As String In SystemObjects.Variables.GetAllNames
                GetSysVars.Add(New SysVar(Variable))
            Next
            Return GetSysVars
        End Function

        ''' <summary>
        ''' Item de As List(Of SysVar)
        ''' </summary>
        ''' <remarks></remarks>
        Public NotInheritable Class SysVar

            ''' <summary>
            ''' Variável
            ''' </summary>
            ''' <remarks></remarks>
            Public Variable As String

            ''' <summary>
            ''' Valor
            ''' </summary>
            ''' <remarks></remarks>
            Public Value As Object

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="Variable">Variável</param>
            ''' <remarks></remarks>
            Public Sub New(Variable As String)
                Me.Variable = Variable
                Me.Value = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable(Variable)
            End Sub

        End Class

    End Class

End Namespace


