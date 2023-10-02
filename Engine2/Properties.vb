'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace Engine2

    ''' <summary>
    ''' Permite a edição de propriedades das entidades gráficas
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Properties

        ''' <summary>
        ''' Edita as propriedades de uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="Color">Cor do AutoCAD</param>
        ''' <param name="LineWeight">Inteiro ou enumerador com a espessura da linha</param>
        ''' <param name="Linetype">Nome do tipo de linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Change(Entity As Entity, Optional Layer As Object = Nothing, Optional Color As Object = Nothing, Optional LineWeight As Object = Nothing, Optional Linetype As Object = Nothing) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim RGB() As String = Nothing
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                        With Entity
                            If IsNothing(Layer) = False Then
                                .Layer = Layer
                            End If
                            If IsNothing(Color) = False Then
                                .Color = Color
                            End If
                            If IsNothing(Linetype) = False Then
                                .Linetype = Linetype
                            End If
                            If IsNothing(LineWeight) = False Then
                                .LineWeight = LineWeight
                            End If
                            .UpgradeOpen()
                        End With
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Properties.Change, motivo: " & ex.Message)
            End Try
        End Function

    End Class
End Namespace


