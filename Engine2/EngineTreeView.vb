Imports System.Windows.Forms

Namespace Engine2

    ''' <summary>
    ''' Carregamento de TreeView
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EngineTreeView

        ''' <summary>
        ''' Carrega uma TreeView com uma array (string)
        ''' </summary>
        ''' <param name="TreeView">Controle TreeView</param>
        ''' <param name="Array">Array com os itens a serem carregados</param>
        ''' <param name="CharSeparator">Caracter separador</param>
        ''' <remarks></remarks>
        Private Sub LoadTreeView(TreeView As System.Windows.Forms.TreeView, Array As String(), CharSeparator As String)
            Dim lastNode As TreeNode = Nothing
            Dim subPathAgg As String
            For Each path As String In Array
                subPathAgg = String.Empty
                For Each subPath As String In path.Split(CharSeparator)
                    subPathAgg += subPath & CharSeparator
                    Dim nodes As TreeNode() = TreeView.Nodes.Find(subPathAgg, True)
                    If nodes.Length = 0 Then
                        If lastNode Is Nothing Then
                            lastNode = TreeView.Nodes.Add(subPathAgg, subPath)
                        Else
                            lastNode = lastNode.Nodes.Add(subPathAgg, subPath)
                        End If
                    Else
                        lastNode = nodes(0)
                    End If
                Next
                lastNode = Nothing
            Next
        End Sub

    End Class

End Namespace


