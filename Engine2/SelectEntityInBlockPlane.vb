Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports System.Linq

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2023
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Permite seleção em blocos ou XRefs
    ''' </summary>
    Public Class SelectEntityInBlockPlane

        ''' <summary>
        ''' Evento de filtro
        ''' </summary>
        ''' <param name="Entity">Entidade a ser avaliada</param>
        ''' <param name="IsValid">Determina se a entidade é válida para seleção</param>
        ''' <param name="CurrentTransaction">Transação corrente</param>
        Public Event SelectionFilter(Entity As Entity, ByRef IsValid As Boolean, CurrentTransaction As Transaction)

        ''' <summary>
        ''' Determina que não foram encontrados itens válidos
        ''' </summary>
        ''' <param name="sender"></param>
        Public Event NotValidate(sender As SelectEntityInBlockPlane)

        ''' <summary>
        ''' Abre o modo de seleção que permite selecionar entidades internas a um bloco
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <returns></returns>
        Public Function GetSelection(BlockReference As BlockReference) As List(Of Entity)

            'Retorno
            GetSelection = New List(Of Entity)

            'Declarações
            Dim TransactionManager As Autodesk.AutoCAD.ApplicationServices.TransactionManager
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim BlockTableRecordSpace As BlockTableRecord
            Dim Clones As List(Of Entity) = Nothing
            Dim Entity As Entity
            Dim Handle As String
            Dim Clone As Entity
            Dim ObjectIdCollection As ObjectIdCollection
            Dim TypedValueCollection As New List(Of TypedValue)
            Dim Result As New List(Of BlockReference)
            Dim ObjectId0 As ObjectId
            Dim ObjectId1 As ObjectId
            Dim [Continue] As Boolean
            Dim IsValid As Boolean
            Dim IsXref As Boolean = False
            Dim ProgressMeter As ProgressMeter = Nothing
            Dim IdCollection As System.Collections.Generic.IEnumerable(Of ObjectId)
            Dim ini As Integer = 0
            Dim OnValidate As Boolean = False

            'Bloqueia o documento
            Using Editor.Document.LockDocument

                'Obtem o gerenciador de transações
                TransactionManager = Database.TransactionManager

                'Cria a transação de interação com o bloco
                Using Transaction0 As Transaction = TransactionManager.StartTransaction

                    Try

                        'Obtem o BlockTableRecord do espaço corrente
                        BlockTableRecordSpace = Transaction0.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)

                        'Obtem o BlockTableRecord do bloco especificado
                        BlockTableRecord = Transaction0.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)

                        'Verifica se é Xref
                        If BlockTableRecord.IsFromExternalReference = True Then
                            IsXref = True
                        Else
                            IsXref = False
                        End If

                        'Mensagem
                        Editor.WriteMessage(vbLf & " -> Abrindo o " & If(IsXref = True, "XRef", "bloco") & " '" & BlockReference.Name & "' para seleção no desenho corrente..." & vbLf)

                        'Instancia a coleção de clones
                        Clones = New List(Of Entity)

                        'Obtem a coleção de BlockTableRecord (Para efeito de contabilizar)
                        IdCollection = BlockTableRecord.Cast(Of ObjectId)

                        'Instancia o temporizador
                        ProgressMeter = New ProgressMeter
                        ProgressMeter.SetLimit(IdCollection.Count)
                        ProgressMeter.Start("Analisando o " & If(IsXref = True, "XRef", "bloco") & ", aguarde...")

                        'Percorre a coleção de entidades do bloco
                        For Each ObjectId0 In IdCollection

                            Try

                                'Obtem a entidade interna do bloco
                                Entity = Transaction0.GetObject(ObjectId0, OpenMode.ForRead)

                                'Verifica se ainda existe
                                If IsNothing(Entity) = False AndAlso Entity.IsErased = False Then

                                    'Habilita para ser considerado
                                    IsValid = True

                                    'Dispara o evento de seleção
                                    RaiseEvent SelectionFilter(Entity, IsValid, Transaction0)

                                    'Avalia se deve ser considerado
                                    If IsValid = True Then

                                        'Cria o clone
                                        Clone = Entity.Clone

                                        'Aplica a matrix de transformação
                                        Clone.TransformBy(BlockReference.BlockTransform)

                                        'Inclui o clone no espaço
                                        BlockTableRecordSpace.AppendEntity(Clone)

                                        'Cria a nova entidade a partir do clone
                                        Transaction0.AddNewlyCreatedDBObject(Clone, True)

                                        'Registra o handle da entidade interna no clone 
                                        Engine2.SetXData(Clone, Transaction0, "CLONEHANDLE", Entity.Handle.ToString)

                                        'Adiciona o clone na coleção
                                        Clones.Add(Clone)

                                        'Seta que houve item válido
                                        OnValidate = True

                                    End If

                                End If

                            Catch

                                'Anula o erro
                                Exit Try

                            End Try

                            'Atualiza o temporizador
                            ProgressMeter.MeterProgress()

                        Next

                        'Finaliza o temporizador
                        ProgressMeter.Stop()

                        'Confirma a transação
                        Transaction0.Commit()

                        'Verifica se existem clones
                        If Clones.Count > 0 Then

                            'Habilita a continuação
                            [Continue] = True

                        Else

                            'Inibe a continuação
                            [Continue] = False

                        End If

                    Catch ex As System.Exception

                        'Aborta a transação
                        Transaction0.Abort()

                        'Erro
                        Throw New System.Exception(ex.Message)

                        'Inibe a continuação
                        [Continue] = False

                    Finally

                        'Avalia se deve continuar
                        If [Continue] = True Then

                            'Monta o filtro de seleção
                            TypedValueCollection = New List(Of TypedValue)
                            With TypedValueCollection
                                .Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "CLONEHANDLE"))
                            End With

                            'Solicita a seleção
                            ObjectIdCollection = Engine2.AcadInterface.GetSelection("Selecione as entidades do bloco '" & BlockReference.Name & "'",, TypedValueCollection)

                            'Avalia
                            If IsNothing(ObjectIdCollection) = False AndAlso ObjectIdCollection.Count > 0 Then

                                'Cria a transação de interação com a seleção
                                Using Transaction1 As Transaction = TransactionManager.StartTransaction

                                    Try

                                        'Obtem o BlockTableRecord do bloco especificado
                                        BlockTableRecord = Transaction1.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)

                                        'Percorre a coleção da seleção
                                        For Each ObjectId1 In ObjectIdCollection

                                            'Abre a entidade para edição
                                            Clone = Transaction1.GetObject(ObjectId1, OpenMode.ForWrite)

                                            'Obtem o handle
                                            Handle = Engine2.XDataEngine.GetXData(Clone, "CLONEHANDLE")

                                            'Percorre a coleção de entidades do bloco
                                            For Each ObjectId0 In BlockTableRecord

                                                Try

                                                    'Obtem a entidade interna do bloco
                                                    Entity = Transaction1.GetObject(ObjectId0, OpenMode.ForRead)

                                                    'Verifica se ainda existe
                                                    If IsNothing(Entity) = False AndAlso Entity.IsErased = False Then

                                                        'Compara o handle
                                                        If Handle = Entity.Handle.ToString Then

                                                            'Atualiza o retorno
                                                            GetSelection.Add(Entity)

                                                        End If

                                                    End If

                                                Catch

                                                    'Anula o erro
                                                    Exit Try

                                                End Try

                                            Next

                                        Next

                                        'Confirma a transação
                                        Transaction1.Commit()

                                    Catch ex As System.Exception

                                        'Aborta a transação
                                        Transaction1.Abort()

                                        'Erro
                                        Throw New System.Exception(ex.Message)

                                    End Try

                                End Using

                            End If

                        End If

                        'Mensagem
                        Editor.WriteMessage(vbLf & " -> Restaurando o " & If(IsXref = True, "XRef", "bloco") & " '" & BlockReference.Name & "'..." & vbLf)

                        'Avalia se existem clones
                        If Clones.Count > 0 Then

                            'Cria a transação 
                            Using Transaction2 As Transaction = TransactionManager.StartTransaction

                                'Percorre a coleção de clones
                                While Clones.Count <> 0

                                    'Obtem o clone
                                    Clone = Transaction2.GetObject(Clones(0).Id, OpenMode.ForWrite)

                                    'Apaga o clone
                                    Clone.Erase()

                                    'Remove o clone da coleção
                                    Clones.Remove(Clone)

                                End While

                                'Conforma a transação
                                Transaction2.Commit()

                            End Using

                        End If

                        'Destroi o temporizador
                        If IsNothing(ProgressMeter) = False Then
                            ProgressMeter.Dispose()
                        End If

                        'Avalia se houve itens válidos
                        If OnValidate = False Then
                            RaiseEvent NotValidate(Me)
                        End If

                    End Try

                End Using

            End Using

            'Retorno
            Return GetSelection

        End Function

    End Class

End Namespace
