'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Engine2

    ''' <summary>
    ''' Classe para gerenciamento de linhas
    ''' </summary>
    ''' <remarks></remarks>
    Public Class LineStyle

        ''' <summary>
        '''  Obtem a coleção de linhas
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        Public Shared Function GetStyles(Transaction As Transaction) As List(Of LinetypeTableRecord)
            GetStyles = New List(Of LinetypeTableRecord)
            Try
                Dim Exceptions As New List(Of String)({"ByLayer", "ByBlock"})
                Dim LinetypeTable As LinetypeTable
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTableRecord As LinetypeTableRecord
                LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                For Each ObjectID As ObjectId In LinetypeTable
                    LinetypeTableRecord = Transaction.GetObject(ObjectID, OpenMode.ForWrite)
                    If Exceptions.Contains(LinetypeTableRecord.Name) = False Then
                        GetStyles.Add(LinetypeTableRecord)
                    End If
                Next
                Return GetStyles
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetStyles, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        '''  Obtem a coleção de linhas
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetStyles() As List(Of LinetypeTableRecord)
            GetStyles = New List(Of LinetypeTableRecord)
            Try
                Dim Exceptions As New List(Of String)({"ByLayer", "ByBlock"})
                Dim LinetypeTable As LinetypeTable
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTableRecord As LinetypeTableRecord
                Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                    LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                    For Each ObjectID As ObjectId In LinetypeTable
                        LinetypeTableRecord = Transaction.GetObject(ObjectID, OpenMode.ForWrite)
                        If Exceptions.Contains(LinetypeTableRecord.Name) = False Then
                            GetStyles.Add(LinetypeTableRecord)
                        End If
                    Next
                End Using
                Return GetStyles
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetStyles, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código '*.Lin' da linha 
        ''' </summary>
        ''' <param name="LinetypeTableRecord">LinetypeTableRecord</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ToLineCode(LinetypeTableRecord As LinetypeTableRecord) As String
            Try
                Dim Dash As String = "A"
                With LinetypeTableRecord
                    For Index As Integer = 0 To .NumDashes - 1
                        Dash = Dash & "," & .DashLengthAt(Index)
                        Try
                            If .TextAt(Index) <> "" Then
                                Dim TextStyleTableRecord As TextStyleTableRecord = Engine2.TextStyle.GetTextStyle(.ShapeStyleAt(Index))
                                If TextStyleTableRecord.IsShapeFile = False Then
                                    Dash = Dash & ",[" & Chr(34) & .TextAt(Index) & Chr(34) & "," & TextStyleTableRecord.Name & ",S=" & .ShapeScaleAt(Index) & ",X=" & .ShapeOffsetAt(Index).X & ",Y=" & .ShapeOffsetAt(Index).Y & ",R=" & .ShapeRotationAt(Index) & "]"
                                Else
                                    MsgBox("A linha contem um ou mais caracteres de um arquivo SHAPE." & vbCrLf & "Os caracteres SHAPE não são suportados pelo sistema e serão omitidos.", MsgBoxStyle.Exclamation)
                                    Exit Try
                                End If
                            End If
                        Catch
                            Exit Try
                        End Try
                    Next
                End With
                Return Dash
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código '*.Lin' da linha 
        ''' </summary>
        ''' <param name="ObjectId">ID do tipo da linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ToLineCode(ByVal ObjectId As ObjectId) As String
            Try
                Dim Dash As String = "A"
                Dim LinetypeTableRecord As LinetypeTableRecord = Engine2.LineStyle.GetLineStyle(ObjectId)
                With LinetypeTableRecord
                    For Index As Integer = 0 To .NumDashes - 1
                        Dash = Dash & "," & .DashLengthAt(Index)
                        Try
                            If .TextAt(Index) <> "" Then
                                Dim TextStyleTableRecord As TextStyleTableRecord = Engine2.TextStyle.GetTextStyle(.ShapeStyleAt(Index))
                                If TextStyleTableRecord.IsShapeFile = False Then
                                    Dash = Dash & ",[" & Chr(34) & .TextAt(Index) & Chr(34) & "," & TextStyleTableRecord.Name & ",S=" & .ShapeScaleAt(Index) & ",X=" & .ShapeOffsetAt(Index).X & ",Y=" & .ShapeOffsetAt(Index).Y & ",R=" & .ShapeRotationAt(Index) & "]"
                                Else
                                    MsgBox("A linha contem um ou mais caracteres de um arquivo SHAPE." & vbCrLf & "Os caracteres SHAPE não são suportados pelo sistema e serão omitidos.", MsgBoxStyle.Exclamation)
                                    Exit Try
                                End If
                            End If
                        Catch
                            Exit Try
                        End Try
                    Next
                End With
                Return If(Dash = "A", "", Dash)
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="LineTypeName">Nome da linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(ByVal LineTypeName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                    LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                    Return LinetypeTable.Has(LineTypeName)
                End Using
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.Contains, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="LineTypeName">Nome da linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(Transaction As Transaction, ByVal LineTypeName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                Return LinetypeTable.Has(LineTypeName)
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.Contains, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="LineTypeName">Nome da linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLineStyle(ByVal LineTypeName As String) As LinetypeTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                Dim LinetypeTableRecord As LinetypeTableRecord
                Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                    LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                    If LinetypeTable.Has(LineTypeName) = True Then
                        LinetypeTableRecord = Transaction.GetObject(LinetypeTable(LineTypeName), OpenMode.ForRead)
                        Return LinetypeTableRecord
                    Else
                        Return Nothing
                    End If
                End Using
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetLineStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LineTypeName">Nome da linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLineStyle(Transaction As Transaction, ByVal LineTypeName As String) As LinetypeTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                Dim LinetypeTableRecord As LinetypeTableRecord
                LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                If LinetypeTable.Has(LineTypeName) = True Then
                    LinetypeTableRecord = Transaction.GetObject(LinetypeTable(LineTypeName), OpenMode.ForRead)
                    Return LinetypeTableRecord
                Else
                    Return Nothing
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetLineStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="ObjectId">ID do tipo de linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLineStyle(ByVal ObjectId As ObjectId) As LinetypeTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                Dim LinetypeTableRecord As LinetypeTableRecord
                Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                    LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                    If LinetypeTable.Has(ObjectId) = True Then
                        LinetypeTableRecord = TryCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), LinetypeTableRecord)
                        Return LinetypeTableRecord
                    Else
                        Return Nothing
                    End If
                End Using
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetLineStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código da linha
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ID do tipo de linha</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLineStyle(Transaction As Transaction, ByVal ObjectId As ObjectId) As LinetypeTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LinetypeTable As LinetypeTable
                Dim LinetypeTableRecord As LinetypeTableRecord
                LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                If LinetypeTable.Has(ObjectId) = True Then
                    LinetypeTableRecord = TryCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), LinetypeTableRecord)
                    Return LinetypeTableRecord
                Else
                    Return Nothing
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.GetLineStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Carrega o tipo de linha
        ''' </summary>
        ''' <param name="Path">Caminho do arquivo de linhas (*.LIN)</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadLineStyle(ByVal Path As String) As Boolean
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Try
                Document.Database.LoadLineTypeFile("*", Path)
                Return True
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                If ex.ErrorStatus = ErrorStatus.FilerError Then
                    Document.Editor.WriteMessage(vbLf & "O arquivo '" & Path & "' não pode ser encontrado")
                    Return False
                ElseIf ex.ErrorStatus = ErrorStatus.DuplicateRecordName Then
                    Return True
                Else
                    Throw New Autodesk.AutoCAD.Runtime.Exception("Erro em LineStyle.LoadLineStyle, motivo: ", ex.Message)
                    Return False
                End If
            End Try
        End Function

        ''' <summary>
        ''' Registra uma linha no AutoCad com base no código da linha (A partir de 'A') 
        ''' </summary>
        ''' <param name="Name">Nome da linha</param>
        ''' <param name="Description">Descrição da linha</param>
        ''' <param name="LineCode">Código da linha</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Name As String, Description As String, LineCode As String) As Boolean
            Try
                'Variaveis
                Dim TextParts As TextParts
                Dim SegmentParts As ArrayList
                Dim TextPart As TextPart
                Dim PatternLength As Double
                Dim NumDashes As Integer
                Dim LinetypeTableRecord As LinetypeTableRecord
                Dim Dash As Object
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                If LineCode.Trim <> "" Then
                    If LineCode.Substring(0, 1).ToUpper <> "A" Then
                        Throw New System.Exception("O código gerador da linha não é válido." & vbCrLf & "O marcador de início 'A' não foi informado.")
                    Else
                        'Obtem os blocos referente as partes que são texto
                        TextParts = New TextParts(LineCode, 1, False)
                        'Trata as partes texto
                        For Index As Integer = 0 To TextParts.Count - 1
                            'Obtem a parte texto
                            TextPart = TextParts.Item(Index)
                            'Cria a fonte
                            Engine2.TextStyle.Create(TextPart.Font.FontFamily.Name, TextPart.Font.FontFamily.Name, 0, TextPart.Font.Bold, TextPart.Font.Italic)
                            'Substitui o bloco de texto no código da linha original acrescido de um caracter 'Alfa' não digitável
                            LineCode = LineCode.Replace(TextPart.Code, Index & Chr(11))
                        Next
                        'Trata as partes não texto 
                        SegmentParts = New ArrayList
                        PatternLength = 0
                        For Each Segment As Object In LineCode.Split(",")
                            'Verifica se é diferente de 'A'
                            If Segment.ToString.ToUpper <> "A" Then
                                'Adiciona o segmento
                                SegmentParts.Add(Segment)
                                'Verifica se o segmento é numérico
                                If Segment.ToString.Contains(Chr(11)) = False Then
                                    If CDbl(Segment) > 0 Then
                                        'Soma o segmento
                                        PatternLength = (PatternLength + CDbl(Segment))
                                    ElseIf CDbl(Segment) < 0 Then
                                        'Reverte o sinal
                                        Segment = (CDbl(Segment) * -1)
                                        'Soma o segmento
                                        PatternLength = (PatternLength + CDbl(Segment))
                                    End If
                                End If
                            End If
                        Next
                        'Obtem o número de partes
                        NumDashes = SegmentParts.Count
                        'Cria a nova linha
                        LinetypeTableRecord = New LinetypeTableRecord
                        With LinetypeTableRecord
                            'Nome da linha
                            .Name = Name
                            'Descrição da linha
                            .AsciiDescription = Description
                            'Comprimento total do segmento
                            .PatternLength = PatternLength
                            'Número de segmentos
                            .NumDashes = NumDashes
                            'Inclusão dos traços
                            For Index As Object = 0 To SegmentParts.Count - 1
                                Dash = SegmentParts.Item(Index)
                                'Verifica se o segmento é numérico
                                If Dash.ToString.Contains(Chr(11)) = False Then
                                    .SetDashLengthAt(Index, CDbl(Dash))
                                Else
                                    'Obtem a parte texto
                                    TextPart = TextParts.Item(Dash.ToString.Substring(0, 1))
                                    'Configura o texto no objeto linha
                                    .SetShapeStyleAt(Index, Engine2.TextStyle.GetTextStyle(TextPart.Font.FontFamily.Name).ObjectId)
                                    .SetShapeNumberAt(Index, If((Index - 1) < 0, 0, (Index - 1)))
                                    .SetShapeOffsetAt(Index, New Vector2d(CDbl(TextPart.Point.X), CDbl(TextPart.Point.Y)))
                                    .SetShapeScaleAt(Index, CDbl(TextPart.Font.Size))
                                    .SetShapeIsUcsOrientedAt(Index, False)
                                    .SetShapeRotationAt(Index, CDbl(TextPart.Angle))
                                    .SetTextAt(Index, TextPart.Text)
                                End If
                            Next
                        End With
                        'Registra a nova linha no AutoCad
                        Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                            'Abre a transação
                            Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                                'Obtem a tabela de linhas do AutoCAD
                                Dim LinetypeTable As LinetypeTable = DirectCast(Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                                'Registra a linha no AutoCAD
                                If LinetypeTable.Has(LinetypeTableRecord.Name) = False Then
                                    LinetypeTable.UpgradeOpen()
                                    LinetypeTable.Add(LinetypeTableRecord)
                                    Transaction.AddNewlyCreatedDBObject(LinetypeTableRecord, True)
                                    'Efetua a transação
                                    Transaction.Commit()
                                End If
                            End Using
                        End Using
                        Return True
                    End If
                Else
                    'Trata as partes não texto 
                    PatternLength = 0
                    'Obtem o número de partes
                    NumDashes = 0
                    'Cria a nova linha
                    LinetypeTableRecord = New LinetypeTableRecord
                    With LinetypeTableRecord
                        'Nome da linha
                        .Name = Name
                        'Descrição da linha
                        .AsciiDescription = Description
                        'Comprimento total do segmento
                        .PatternLength = PatternLength
                        'Número de segmentos
                        .NumDashes = NumDashes
                    End With
                    'Registra a nova linha no AutoCad
                    Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        'Abre a transação
                        Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                            'Obtem a tabela de linhas do AutoCAD
                            Dim LinetypeTable As LinetypeTable = DirectCast(Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                            'Registra a linha no AutoCAD
                            Dim ObjectId As ObjectId
                            If LinetypeTable.Has(LinetypeTableRecord.Name) = False Then
                                LinetypeTable.UpgradeOpen()
                                ObjectId = LinetypeTable.Add(LinetypeTableRecord)
                                Transaction.AddNewlyCreatedDBObject(LinetypeTableRecord, True)
                                'Efetua a transação
                                Transaction.Commit()
                            End If
                        End Using
                    End Using
                    Return True
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.Create, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Registra uma linha no AutoCad com base no código da linha (A partir de 'A') 
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Name">Nome da linha</param>
        ''' <param name="Description">Descrição da linha</param>
        ''' <param name="LineCode">Código da linha</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Transaction As Transaction, Name As String, Description As String, LineCode As String) As Boolean
            Try
                'Variaveis
                Dim TextParts As TextParts
                Dim SegmentParts As ArrayList
                Dim TextPart As TextPart
                Dim PatternLength As Double
                Dim NumDashes As Integer
                Dim LinetypeTableRecord As LinetypeTableRecord
                Dim Dash As Object
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                If LineCode.Trim <> "" Then
                    If LineCode.Substring(0, 1).ToUpper <> "A" Then
                        Throw New System.Exception("O código gerador da linha não é válido." & vbCrLf & "O marcador de início 'A' não foi informado.")
                    Else
                        'Obtem os blocos referente as partes que são texto
                        TextParts = New TextParts(LineCode, 1, False)
                        'Trata as partes texto
                        For Index As Integer = 0 To TextParts.Count - 1
                            'Obtem a parte texto
                            TextPart = TextParts.Item(Index)
                            'Cria a fonte
                            Engine2.TextStyle.Create(TextPart.Font.FontFamily.Name, TextPart.Font.FontFamily.Name, 0, TextPart.Font.Bold, TextPart.Font.Italic)
                            'Substitui o bloco de texto no código da linha original acrescido de um caracter 'Alfa' não digitável
                            LineCode = LineCode.Replace(TextPart.Code, Index & Chr(11))
                        Next
                        'Trata as partes não texto 
                        SegmentParts = New ArrayList
                        PatternLength = 0
                        For Each Segment As Object In LineCode.Split(",")
                            'Verifica se é diferente de 'A'
                            If Segment.ToString.ToUpper <> "A" Then
                                'Adiciona o segmento
                                SegmentParts.Add(Segment)
                                'Verifica se o segmento é numérico
                                If Segment.ToString.Contains(Chr(11)) = False Then
                                    If CDbl(Segment) > 0 Then
                                        'Soma o segmento
                                        PatternLength = (PatternLength + CDbl(Segment))
                                    ElseIf CDbl(Segment) < 0 Then
                                        'Reverte o sinal
                                        Segment = (CDbl(Segment) * -1)
                                        'Soma o segmento
                                        PatternLength = (PatternLength + CDbl(Segment))
                                    End If
                                End If
                            End If
                        Next
                        'Obtem o número de partes
                        NumDashes = SegmentParts.Count
                        'Cria a nova linha
                        LinetypeTableRecord = New LinetypeTableRecord
                        With LinetypeTableRecord
                            'Nome da linha
                            .Name = Name
                            'Descrição da linha
                            .AsciiDescription = Description
                            'Comprimento total do segmento
                            .PatternLength = PatternLength
                            'Número de segmentos
                            .NumDashes = NumDashes
                            'Inclusão dos traços
                            For Index As Object = 0 To SegmentParts.Count - 1
                                Dash = SegmentParts.Item(Index)
                                'Verifica se o segmento é numérico
                                If Dash.ToString.Contains(Chr(11)) = False Then
                                    .SetDashLengthAt(Index, CDbl(Dash))
                                Else
                                    'Obtem a parte texto
                                    TextPart = TextParts.Item(Dash.ToString.Substring(0, 1))
                                    'Configura o texto no objeto linha
                                    .SetShapeStyleAt(Index, Engine2.TextStyle.GetTextStyle(TextPart.Font.FontFamily.Name).ObjectId)
                                    .SetShapeNumberAt(Index, If((Index - 1) < 0, 0, (Index - 1)))
                                    .SetShapeOffsetAt(Index, New Vector2d(CDbl(TextPart.Point.X), CDbl(TextPart.Point.Y)))
                                    .SetShapeScaleAt(Index, CDbl(TextPart.Font.Size))
                                    .SetShapeIsUcsOrientedAt(Index, False)
                                    .SetShapeRotationAt(Index, CDbl(TextPart.Angle))
                                    .SetTextAt(Index, TextPart.Text)
                                End If
                            Next
                        End With
                        'Obtem a tabela de linhas do AutoCAD
                        Dim LinetypeTable As LinetypeTable = DirectCast(Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                        'Registra a linha no AutoCAD
                        If LinetypeTable.Has(LinetypeTableRecord.Name) = False Then
                            LinetypeTable.UpgradeOpen()
                            LinetypeTable.Add(LinetypeTableRecord)
                            Transaction.AddNewlyCreatedDBObject(LinetypeTableRecord, True)
                        End If
                        Return True
                    End If
                Else
                    'Trata as partes não texto 
                    PatternLength = 0
                    'Obtem o número de partes
                    NumDashes = 0
                    'Cria a nova linha
                    LinetypeTableRecord = New LinetypeTableRecord
                    With LinetypeTableRecord
                        'Nome da linha
                        .Name = Name
                        'Descrição da linha
                        .AsciiDescription = Description
                        'Comprimento total do segmento
                        .PatternLength = PatternLength
                        'Número de segmentos
                        .NumDashes = NumDashes
                    End With
                    'Obtem a tabela de linhas do AutoCAD
                    Dim LinetypeTable As LinetypeTable = DirectCast(Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                    'Registra a linha no AutoCAD
                    Dim ObjectId As ObjectId
                    If LinetypeTable.Has(LinetypeTableRecord.Name) = False Then
                        LinetypeTable.UpgradeOpen()
                        ObjectId = LinetypeTable.Add(LinetypeTableRecord)
                        Transaction.AddNewlyCreatedDBObject(LinetypeTableRecord, True)
                    End If
                    Return True
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineStyle.Create, motivo: " & ex.Message)
            End Try
        End Function

    End Class

End Namespace
