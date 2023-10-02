Imports System.Text

Namespace Engine2

    '=========================================================================================================='
    'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
    'EM:2014
    'PARA:SPHE ENGENHARIA
    '=========================================================================================================='
    'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
    '=========================================================================================================='

    ''' <summary>
    ''' Permite melhor monitoramento dos erros
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Debug

        ''' <summary>
        ''' Obtem a linha com o erro 
        ''' </summary>
        ''' <param name="Ex">Exception</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLineNumberError(ByVal Ex As System.Exception) As Integer
            Dim lineNumber As Int32 = 0
            Dim lineSearch As String
            Dim Index As Integer
            lineSearch = ":line "
            Index = Ex.StackTrace.LastIndexOf(lineSearch)
            If Index <> -1 Then
                GetLineNumberError = Ex.StackTrace.Substring(Index + lineSearch.Length)
            Else
                GetLineNumberError = -1
            End If
            Return GetLineNumberError
        End Function

        ''' <summary>
        ''' Retorna a mensagem de erro
        ''' </summary>
        ''' <param name="Ex">Exception</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetExceptionMessage(Ex As System.Exception) As String
            Dim Items() As String
            Dim ValidValues As New List(Of String)
            Dim SB As New StringBuilder
            Dim lineNumber As Int32 = 0
            Dim lineSearch As String
            Dim Index As Integer
            Dim Local As String
            lineSearch = "\"
            Index = Ex.StackTrace.LastIndexOf(lineSearch)
            Local = Ex.StackTrace.Substring(Index + lineSearch.Length)
            GetExceptionMessage = " " & Ex.StackTrace
            GetExceptionMessage = Ex.StackTrace.Replace(" at ", Chr(11))
            GetExceptionMessage = GetExceptionMessage.Replace(" in ", Chr(11))
            Items = GetExceptionMessage.Split(Chr(11))
            For Each Value As String In Items
                If Value.Trim <> "" And Value.Trim <> Ex.Message Then
                    ValidValues.Add(Value.Trim)
                End If
            Next
            With SB
                .AppendLine("Erro: " & Local)
                If Ex.Message.Trim <> "" Then
                    .AppendLine("")
                    .AppendLine("MENSAGEM")
                    .AppendLine(Ex.Message)
                End If
                If ValidValues.Count > 0 Then
                    .AppendLine("")
                    .AppendLine("INFORMAÇÕES DO DEPURADOR")
                    For Each Value As String In ValidValues
                        .AppendLine("-> " & Value)
                    Next
                End If
            End With
            Return SB.ToString
        End Function

    End Class

End Namespace

