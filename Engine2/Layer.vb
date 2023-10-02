'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Reflection
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.Colors

Namespace Engine2

    ''' <summary>
    ''' Classe para gerenciamento de layers
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Layer

        ''' <summary>
        ''' Tipos de bloqueio na camada
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eLockType
            Locked = 2
            Frozen = 4
            Off = 8
            Hidden = 16
            All = 32
        End Enum

        ''' <summary>
        ''' Enumera as cores
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eColors
            Red = 1
            Yellow = 2
            Green = 3
            Cyan = 4
            Blue = 5
            Magenta = 6
            White = 7
        End Enum

        ''' <summary>
        ''' Converte a entidade para o layer determinado
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function ToLayer(ByVal Entity As Entity, ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable
                Dim LayerTableRecord As LayerTableRecord
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                        LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                        With LayerTableRecord
                            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                            Entity.Layer = LayerTableRecord.Name
                            Entity.Color = .Color
                            Entity.LinetypeId = .LinetypeObjectId
                            Entity.LineWeight = .LineWeight
                            .UpgradeOpen()
                        End With
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.ConvertToLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Converte a entidade para o layer determinado
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function ToLayer(Transaction As Transaction, ByVal Entity As Entity, ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable
                Dim LayerTableRecord As LayerTableRecord
                LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                With LayerTableRecord
                    Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                    Entity.Layer = LayerTableRecord.Name
                    Entity.Color = .Color
                    Entity.LinetypeId = .LinetypeObjectId
                    Entity.LineWeight = .LineWeight
                    .UpgradeOpen()
                End With
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.ConvertToLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem a entidade layer
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function GetLayer(ByVal LayerName As String) As LayerTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable
                Dim LayerTableRecord As LayerTableRecord
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                        If LayerTable.Has(LayerName) = True Then
                            LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                            Return LayerTableRecord
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem a entidade layer
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function GetLayer(Transaction As Transaction, ByVal LayerName As String) As LayerTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable
                Dim LayerTableRecord As LayerTableRecord
                LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                If LayerTable.Has(LayerName) = True Then
                    LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                    Return LayerTableRecord
                Else
                    Return Nothing
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Verifica se o layer existe
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                        Return LayerTable.Has(LayerName)
                        Transaction.Dispose()
                    End Using
                End Using
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.LayerContains, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Verifica se o layer existe
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(Transaction As Transaction, ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                Return LayerTable.Has(LayerName)
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.LayerContains, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Modifica os modos de acesso a um layer
        ''' </summary>
        ''' <param name="LayerName">Nome da camada (Aceita filtro)</param>
        ''' <param name="IsLocked">Bloquear\Desbloquear</param>
        ''' <param name="IsFrozen">Congelar\Descongelar</param>
        ''' <param name="IsPlottable">Imprimível\Não imprimível</param>
        ''' <param name="ColorIndex">Número da cor</param>
        ''' <param name="LinetypeName">Nome do tipo de linha</param>
        ''' <param name="LineWeight">Espessura da linha (Autodesk.AutoCAD.DatabaseServices.LineWeight)</param>
        ''' <param name="IsOff">Determina se a camada é visível</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChangeAccess(ByVal LayerName As String, ByVal IsLocked As Object, ByVal IsFrozen As Object, Optional ByVal IsPlottable As Object = Nothing, Optional ByVal ColorIndex As Object = Nothing, Optional ByVal LineWeight As Object = Nothing, Optional ByVal LinetypeName As Object = Nothing, Optional IsOff As Object = Nothing) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                If CurrentLayer = LayerName And IsOff = True Then
                    For Each Layer As String In Engine2.Layer.GetNames(True)
                        If Layer <> CurrentLayer And Layer.Contains("|") = False Then
                            If Engine2.Layer.IsLocked(Layer, eLockType.Off) = False Then
                                Engine2.Layer.SetCurrent(Layer)
                                Exit For
                            End If
                        End If
                    Next
                End If
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing
                Dim FilterLayers As ArrayList = Nothing
                Dim LinetypeTable As LinetypeTable
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                        If LayerTable.Has(LayerName) = True Then
                            LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                            With LayerTableRecord
                                If IsNothing(IsLocked) = False Then
                                    .IsLocked = IsLocked
                                End If
                                If IsNothing(IsFrozen) = False Then
                                    .IsFrozen = IsFrozen
                                End If
                                If IsNothing(IsPlottable) = False Then
                                    .IsPlottable = IsPlottable
                                End If
                                If IsNothing(ColorIndex) = False Then
                                    .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                                End If
                                If IsNothing(LineWeight) = False Then
                                    .LineWeight = LineWeight
                                End If
                                If IsNothing(IsOff) = False Then
                                    If IsOff = True Then
                                        If CurrentLayer <> LayerName Then
                                            .IsOff = IsOff
                                        Else
                                            MsgBox("A camada '" & LayerName & "' é a camada corrente e não pode ser desligada.", MsgBoxStyle.Exclamation)
                                        End If
                                    Else
                                        .IsOff = IsOff
                                    End If
                                End If
                                If IsNothing(LinetypeName) = False Then
                                    LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                                    With LinetypeTable
                                        If .Has(LinetypeName) = True Then
                                            LayerTableRecord.UpgradeOpen()
                                            LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                                        End If
                                    End With
                                End If
                                .UpgradeOpen()
                            End With
                            Transaction.Commit()
                            Document.Editor.WriteMessage(vbLf & "Layer '" & LayerName & "' modificado com sucesso." & vbLf)
                        Else
                            FilterLayers = Engine2.Layer.GetNames(True, New ArrayList({LayerName}))
                            For Each LayerName In FilterLayers
                                LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                                With LayerTableRecord
                                    If IsNothing(IsLocked) = False Then
                                        .IsLocked = IsLocked
                                    End If
                                    If IsNothing(IsFrozen) = False Then
                                        .IsFrozen = IsFrozen
                                    End If
                                    If IsNothing(IsPlottable) = False Then
                                        .IsPlottable = IsPlottable
                                    End If
                                    If IsNothing(ColorIndex) = False Then
                                        .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                                    End If
                                    If IsNothing(LineWeight) = False Then
                                        .LineWeight = LineWeight
                                    End If
                                    If IsNothing(IsOff) = False Then
                                        If IsOff = True Then
                                            If CurrentLayer <> LayerName Then
                                                .IsOff = IsOff
                                            Else
                                                MsgBox("A camada '" & LayerName & "' é a camada corrente e não pode ser desligada.", MsgBoxStyle.Exclamation)
                                            End If
                                        Else
                                            .IsOff = IsOff
                                        End If
                                    End If
                                    If IsNothing(LinetypeName) = False Then
                                        LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                                        With LinetypeTable
                                            If .Has(LinetypeName) = True Then
                                                LayerTableRecord.UpgradeOpen()
                                                LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                                            End If
                                        End With
                                    End If
                                    .UpgradeOpen()
                                End With
                                Document.Editor.WriteMessage(vbLf & "Layer '" & LayerName & "' modificado com sucesso." & vbLf)
                            Next
                            Transaction.Commit()
                        End If
                    End Using
                End Using
                Return True
            Catch
                Return False
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Modifica os modos de acesso a um layer
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome da camada (Aceita filtro)</param>
        ''' <param name="IsLocked">Bloquear\Desbloquear</param>
        ''' <param name="IsFrozen">Congelar\Descongelar</param>
        ''' <param name="IsPlottable">Imprimível\Não imprimível</param>
        ''' <param name="ColorIndex">Número da cor</param>
        ''' <param name="LinetypeName">Nome do tipo de linha</param>
        ''' <param name="LineWeight">Espessura da linha (Autodesk.AutoCAD.DatabaseServices.LineWeight)</param>
        ''' <param name="IsOff">Determina se a camada é visível</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChangeAccess(Transaction As Transaction, ByVal LayerName As String, ByVal IsLocked As Object, ByVal IsFrozen As Object, Optional ByVal IsPlottable As Object = Nothing, Optional ByVal ColorIndex As Object = Nothing, Optional ByVal LineWeight As Object = Nothing, Optional ByVal LinetypeName As Object = Nothing, Optional IsOff As Object = Nothing) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                If CurrentLayer = LayerName And IsOff = True Then
                    For Each Layer As String In Engine2.Layer.GetNames(True)
                        If Layer <> CurrentLayer And Layer.Contains("|") = False Then
                            If Engine2.Layer.IsLocked(Layer, eLockType.Off) = False Then
                                Engine2.Layer.SetCurrent(Layer)
                                Exit For
                            End If
                        End If
                    Next
                End If
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing
                Dim FilterLayers As ArrayList = Nothing
                Dim LinetypeTable As LinetypeTable
                Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                If LayerTable.Has(LayerName) = True Then
                    LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                    With LayerTableRecord
                        If IsNothing(IsLocked) = False Then
                            .IsLocked = IsLocked
                        End If
                        If IsNothing(IsFrozen) = False Then
                            .IsFrozen = IsFrozen
                        End If
                        If IsNothing(IsPlottable) = False Then
                            .IsPlottable = IsPlottable
                        End If
                        If IsNothing(ColorIndex) = False Then
                            .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                        End If
                        If IsNothing(LineWeight) = False Then
                            .LineWeight = LineWeight
                        End If
                        If IsNothing(IsOff) = False Then
                            If IsOff = True Then
                                If CurrentLayer <> LayerName Then
                                    .IsOff = IsOff
                                Else
                                    MsgBox("A camada '" & LayerName & "' é a camada corrente e não pode ser desligada.", MsgBoxStyle.Exclamation)
                                End If
                            Else
                                .IsOff = IsOff
                            End If
                        End If
                        If IsNothing(LinetypeName) = False Then
                            LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                            With LinetypeTable
                                If .Has(LinetypeName) = True Then
                                    LayerTableRecord.UpgradeOpen()
                                    LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                                End If
                            End With
                        End If
                        .UpgradeOpen()
                    End With
                    Document.Editor.WriteMessage(vbLf & "Layer '" & LayerName & "' modificado com sucesso." & vbLf)
                Else
                    FilterLayers = Engine2.Layer.GetNames(True, New ArrayList({LayerName}))
                    For Each LayerName In FilterLayers
                        LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                        With LayerTableRecord
                            If IsNothing(IsLocked) = False Then
                                .IsLocked = IsLocked
                            End If
                            If IsNothing(IsFrozen) = False Then
                                .IsFrozen = IsFrozen
                            End If
                            If IsNothing(IsPlottable) = False Then
                                .IsPlottable = IsPlottable
                            End If
                            If IsNothing(ColorIndex) = False Then
                                .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                            End If
                            If IsNothing(LineWeight) = False Then
                                .LineWeight = LineWeight
                            End If
                            If IsNothing(IsOff) = False Then
                                If IsOff = True Then
                                    If CurrentLayer <> LayerName Then
                                        .IsOff = IsOff
                                    Else
                                        MsgBox("A camada '" & LayerName & "' é a camada corrente e não pode ser desligada.", MsgBoxStyle.Exclamation)
                                    End If
                                Else
                                    .IsOff = IsOff
                                End If
                            End If
                            If IsNothing(LinetypeName) = False Then
                                LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                                With LinetypeTable
                                    If .Has(LinetypeName) = True Then
                                        LayerTableRecord.UpgradeOpen()
                                        LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                                    End If
                                End With
                            End If
                            .UpgradeOpen()
                        End With
                        Document.Editor.WriteMessage(vbLf & "Layer '" & LayerName & "' modificado com sucesso." & vbLf)
                    Next
                End If
                Return True
            Catch
                Return False
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Criar layer
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <param name="LinetypeName">Tipo de linha</param>
        ''' <param name="IsLocked">Bloqueio de layer</param>
        ''' <param name="IsFrozen">Congelamento de layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(ByVal LayerName As String, ByVal ColorIndex As eColors, Optional ByVal LineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = DatabaseServices.LineWeight.LineWeight005, Optional ByVal LinetypeName As String = "Continuous", Optional ByVal IsLocked As Boolean = False, Optional ByVal IsFrozen As Boolean = False, Optional ByVal IsPlottable As Boolean = True) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                        If LayerTable.Has(LayerName) = False Then
                            LayerTableRecord = New LayerTableRecord()
                            With LayerTableRecord
                                .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                                .Name = LayerName
                                .IsLocked = IsLocked
                                .IsFrozen = IsFrozen
                                .LineWeight = LineWeight
                                .IsPlottable = IsPlottable
                            End With
                            LayerTable.Add(LayerTableRecord)
                            Transaction.AddNewlyCreatedDBObject(LayerTableRecord, True)
                        Else
                            Document.Database.Clayer = LayerTable("0")
                            LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                            'If Update = True Then
                            With LayerTableRecord
                                .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                                .Name = LayerName
                                .IsLocked = IsLocked
                                .IsFrozen = IsFrozen
                                .LineWeight = LineWeight
                                .IsPlottable = IsPlottable
                                .UpgradeOpen()
                            End With
                            'End If
                        End If
                        Document.Database.Clayer = LayerTable(LayerName)
                        Dim LinetypeTable As LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                        With LinetypeTable
                            If .Has(LinetypeName) = True Then
                                LayerTableRecord.UpgradeOpen()
                                LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                            End If
                        End With
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.Create, motivo: " & ex.Message)
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Criar layer
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <param name="LinetypeName">Tipo de linha</param>
        ''' <param name="IsLocked">Bloqueio de layer</param>
        ''' <param name="IsFrozen">Congelamento de layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Transaction As Transaction, ByVal LayerName As String, ByVal ColorIndex As eColors, Optional ByVal LineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = DatabaseServices.LineWeight.LineWeight005, Optional ByVal LinetypeName As String = "Continuous", Optional ByVal IsLocked As Boolean = False, Optional ByVal IsFrozen As Boolean = False, Optional ByVal IsPlottable As Boolean = True) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing
                Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                If LayerTable.Has(LayerName) = False Then
                    LayerTableRecord = New LayerTableRecord()
                    With LayerTableRecord
                        .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                        .Name = LayerName
                        .IsLocked = IsLocked
                        .IsFrozen = IsFrozen
                        .LineWeight = LineWeight
                        .IsPlottable = IsPlottable
                    End With
                    LayerTable.Add(LayerTableRecord)
                    Transaction.AddNewlyCreatedDBObject(LayerTableRecord, True)
                Else
                    Document.Database.Clayer = LayerTable("0")
                    LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                    'If Update = True Then
                    With LayerTableRecord
                        .Color = Color.FromColorIndex(ColorMethod.ByAci, ColorIndex)
                        .Name = LayerName
                        .IsLocked = IsLocked
                        .IsFrozen = IsFrozen
                        .LineWeight = LineWeight
                        .IsPlottable = IsPlottable
                        .UpgradeOpen()
                    End With
                    'End If
                End If
                Document.Database.Clayer = LayerTable(LayerName)
                Dim LinetypeTable As LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                With LinetypeTable
                    If .Has(LinetypeName) = True Then
                        LayerTableRecord.UpgradeOpen()
                        LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                    End If
                End With
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.Create, motivo: " & ex.Message)
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Criar layer
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="Color">Cor</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <param name="LinetypeName">Tipo de linha</param>
        ''' <param name="IsLocked">Bloqueio de layer</param>
        ''' <param name="IsFrozen">Congelamento de layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(ByVal LayerName As String, ByVal Color As Autodesk.AutoCAD.Colors.Color, Optional ByVal LineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = DatabaseServices.LineWeight.LineWeight005, Optional ByVal LinetypeName As String = "Continuous", Optional ByVal IsLocked As Boolean = False, Optional ByVal IsFrozen As Boolean = False, Optional ByVal IsPlottable As Boolean = True) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                Dim _Color As Color
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                        If LayerTable.Has(LayerName) = False Then
                            LayerTableRecord = New LayerTableRecord()
                            With LayerTableRecord
                                _Color = Color.FromColorIndex(Colors.ColorMethod.ByAci, Color.ColorIndex)
                                If _Color.ColorIndex > 0 And _Color.ColorIndex < 256 Then
                                    .Color = _Color
                                Else
                                    .Color = Color
                                End If
                                .Name = LayerName
                                .IsLocked = IsLocked
                                .IsFrozen = IsFrozen
                                .LineWeight = LineWeight
                                .IsPlottable = IsPlottable
                            End With
                            LayerTable.Add(LayerTableRecord)
                            Transaction.AddNewlyCreatedDBObject(LayerTableRecord, True)
                        Else
                            Document.Database.Clayer = LayerTable("0")
                            LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                            With LayerTableRecord
                                _Color = Color.FromColorIndex(Colors.ColorMethod.ByAci, Color.ColorIndex)
                                If _Color.ColorIndex > 0 And _Color.ColorIndex < 256 Then
                                    .Color = _Color
                                Else
                                    .Color = Color
                                End If
                                .Name = LayerName
                                .IsLocked = IsLocked
                                .IsFrozen = IsFrozen
                                .LineWeight = LineWeight
                                .IsPlottable = IsPlottable
                                .UpgradeOpen()
                            End With
                        End If
                        Document.Database.Clayer = LayerTable(LayerName)
                        Dim LinetypeTable As LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                        With LinetypeTable
                            If .Has(LinetypeName) = True Then
                                LayerTableRecord.UpgradeOpen()
                                LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                            Else
                            End If
                        End With
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.Create, motivo: " & ex.Message)
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Criar layer
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="Color">Cor</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <param name="LinetypeName">Tipo de linha</param>
        ''' <param name="IsLocked">Bloqueio de layer</param>
        ''' <param name="IsFrozen">Congelamento de layer</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Transaction As Transaction, ByVal LayerName As String, ByVal Color As Autodesk.AutoCAD.Colors.Color, Optional ByVal LineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = DatabaseServices.LineWeight.LineWeight005, Optional ByVal LinetypeName As String = "Continuous", Optional ByVal IsLocked As Boolean = False, Optional ByVal IsFrozen As Boolean = False, Optional ByVal IsPlottable As Boolean = True) As Boolean
            Dim CurrentLayer As String = Engine2.Layer.GetCurrent
            Try
                Dim _Color As Color
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord = Nothing

                Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                If LayerTable.Has(LayerName) = False Then
                    LayerTableRecord = New LayerTableRecord()
                    With LayerTableRecord
                        _Color = Color.FromColorIndex(Colors.ColorMethod.ByAci, Color.ColorIndex)
                        If _Color.ColorIndex > 0 And _Color.ColorIndex < 256 Then
                            .Color = _Color
                        Else
                            .Color = Color
                        End If
                        .Name = LayerName
                        .IsLocked = IsLocked
                        .IsFrozen = IsFrozen
                        .LineWeight = LineWeight
                        .IsPlottable = IsPlottable
                    End With
                    LayerTable.Add(LayerTableRecord)
                    Transaction.AddNewlyCreatedDBObject(LayerTableRecord, True)
                Else
                    Document.Database.Clayer = LayerTable("0")
                    LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForWrite)
                    With LayerTableRecord
                        _Color = Color.FromColorIndex(Colors.ColorMethod.ByAci, Color.ColorIndex)
                        If _Color.ColorIndex > 0 And _Color.ColorIndex < 256 Then
                            .Color = _Color
                        Else
                            .Color = Color
                        End If
                        .Name = LayerName
                        .IsLocked = IsLocked
                        .IsFrozen = IsFrozen
                        .LineWeight = LineWeight
                        .IsPlottable = IsPlottable
                        .UpgradeOpen()
                    End With
                End If
                Document.Database.Clayer = LayerTable(LayerName)
                Dim LinetypeTable As LinetypeTable = Transaction.GetObject(Document.Database.LinetypeTableId, OpenMode.ForRead)
                With LinetypeTable
                    If .Has(LinetypeName) = True Then
                        LayerTableRecord.UpgradeOpen()
                        LayerTableRecord.LinetypeObjectId = LinetypeTable(LinetypeName)
                    Else
                    End If
                End With
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.Create, motivo: " & ex.Message)
            Finally
                Engine2.Layer.SetCurrent(CurrentLayer)
            End Try
        End Function

        ''' <summary>
        ''' Seta o layer corrente
        ''' </summary>
        ''' <param name="LayerName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SetCurrent(ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                        Document.Database.Clayer = LayerTable(LayerName)
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.SetLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Seta o layer corrente
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome da camada</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SetCurrent(Transaction As Transaction, ByVal LayerName As String) As Boolean
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForWrite)
                Document.Database.Clayer = LayerTable(LayerName)
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.SetLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o Layer corrente
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrent() As String
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        LayerTableRecord = CType(Transaction.GetObject(Document.Database.Clayer, OpenMode.ForRead), LayerTableRecord)
                    End Using
                End Using
                Return LayerTableRecord.Name
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CurrentLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o Layer corrente
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrent(Transaction As Transaction) As String
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord
                LayerTableRecord = CType(Transaction.GetObject(Document.Database.Clayer, OpenMode.ForRead), LayerTableRecord)
                Return LayerTableRecord.Name
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CurrentLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Analisa se o layer é editável
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsLocked(LayerName As String) As Object
            Try
                Dim Msg As String = ""
                If LayerName.Trim <> "" Then
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using DocumentLock As DocumentLock = Document.LockDocument()
                        Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                            Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                            Dim LayerTableRecord As LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                            With LayerTableRecord
                                If .IsLocked = True Then
                                    Msg = "camada '" & LayerName & "' esta bloqueada."
                                End If
                                If .IsFrozen = True Then
                                    Msg = "camada '" & LayerName & "' esta congelada."
                                End If
                                If .IsOff = True Then
                                    Msg = "camada '" & LayerName & "' esta desligada."
                                End If
                                If .IsHidden = True Then
                                    Msg = "camada '" & LayerName & "' esta oculta."
                                End If
                            End With
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                If Msg.Trim <> "" Then
                    Return Msg
                Else
                    Return Nothing
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CheckLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Analisa se o layer é editável
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsLocked(Transaction As Transaction, LayerName As String) As Object
            Try
                Dim Msg As String = ""
                If LayerName.Trim <> "" Then
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                    Dim LayerTableRecord As LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                    With LayerTableRecord
                        If .IsLocked = True Then
                            Msg = "camada '" & LayerName & "' esta bloqueada."
                        End If
                        If .IsFrozen = True Then
                            Msg = "camada '" & LayerName & "' esta congelada."
                        End If
                        If .IsOff = True Then
                            Msg = "camada '" & LayerName & "' esta desligada."
                        End If
                        If .IsHidden = True Then
                            Msg = "camada '" & LayerName & "' esta oculta."
                        End If
                    End With
                End If
                If Msg.Trim <> "" Then
                    Return Msg
                Else
                    Return Nothing
                End If
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CheckLayer, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Analisa se o layer é editável
        ''' </summary>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="LockType">Tipo de bloqueio a ser verificado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsLocked(LayerName As String, Optional LockType As eLockType = eLockType.All) As Boolean
            Try
                Dim Msg As String = ""
                If LayerName.Trim <> "" Then
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using DocumentLock As DocumentLock = Document.LockDocument()
                        Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                            Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                            Dim LayerTableRecord As LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                            With LayerTableRecord
                                Select Case LockType
                                    Case eLockType.All
                                        If .IsLocked = True Then
                                            Return True
                                        End If
                                        If .IsFrozen = True Then
                                            Return True
                                        End If
                                        If .IsOff = True Then
                                            Return True
                                        End If
                                        If .IsHidden = True Then
                                            Return True
                                        End If
                                        Return False
                                    Case eLockType.Frozen
                                        If .IsFrozen = True Then
                                            Return True
                                        End If
                                        Return False
                                    Case eLockType.Hidden
                                        If .IsHidden = True Then
                                            Return True
                                        End If
                                        Return False
                                    Case eLockType.Locked
                                        If .IsLocked = True Then
                                            Return True
                                        End If
                                        Return False
                                    Case eLockType.Off
                                        If .IsOff = True Then
                                            Return True
                                        End If
                                        Return False
                                End Select
                            End With
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return False
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CheckLayerll, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Analisa se o layer é editável
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="LayerName">Nome do layer</param>
        ''' <param name="LockType">Tipo de bloqueio a ser verificado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsLocked(Transaction As Transaction, LayerName As String, Optional LockType As eLockType = eLockType.All) As Boolean
            Try
                Dim Msg As String = ""
                If LayerName.Trim <> "" Then
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Dim LayerTable As LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                    Dim LayerTableRecord As LayerTableRecord = Transaction.GetObject(LayerTable(LayerName), OpenMode.ForRead)
                    With LayerTableRecord
                        Select Case LockType
                            Case eLockType.All
                                If .IsLocked = True Then
                                    Return True
                                End If
                                If .IsFrozen = True Then
                                    Return True
                                End If
                                If .IsOff = True Then
                                    Return True
                                End If
                                If .IsHidden = True Then
                                    Return True
                                End If
                                Return False
                            Case eLockType.Frozen
                                If .IsFrozen = True Then
                                    Return True
                                End If
                                Return False
                            Case eLockType.Hidden
                                If .IsHidden = True Then
                                    Return True
                                End If
                                Return False
                            Case eLockType.Locked
                                If .IsLocked = True Then
                                    Return True
                                End If
                                Return False
                            Case eLockType.Off
                                If .IsOff = True Then
                                    Return True
                                End If
                                Return False
                        End Select
                    End With
                End If
                Return False
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.CheckLayerll, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna a coleção de layers
        ''' </summary>
        '''<param name="IncludeLayerZero">Define se o layer zero será incluído</param>
        '''<param name="Filters">Filtros para seleção de camadas</param>
        ''' <param name="IsFrozen">Determina se camadas congeladas serão aceitas</param>
        ''' <param name="IsLocked">Determina se camadas bloqueadas serão aceitas</param>
        '''<returns>Boolean</returns>
        '''<remarks></remarks>
        Public Shared Function GetNames(Optional IncludeLayerZero As Boolean = False, Optional Filters As ArrayList = Nothing, Optional IsLocked As Object = Nothing, Optional IsFrozen As Object = Nothing) As ArrayList
            Try
                GetNames = New ArrayList
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord
                Dim LayerTable As LayerTable
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                        For Each ObjectID As ObjectId In LayerTable
                            LayerTableRecord = DirectCast(Transaction.GetObject(ObjectID, OpenMode.ForWrite), LayerTableRecord)
                            If IsNothing(Filters) = False Then
                                For Each Filter As String In Filters
                                    If LayerTableRecord.Name Like Filter = True Then
                                        If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                            GetNames.Add(LayerTableRecord.Name)
                                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                            If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                            If LayerTableRecord.IsFrozen = IsFrozen Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                            If LayerTableRecord.IsLocked = IsLocked Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                    GetNames.Add(LayerTableRecord.Name)
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                    If LayerTableRecord.IsLocked = IsLocked Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                End If
                            End If
                        Next
                    End Using
                End Using
                If IncludeLayerZero = False Then
                    If GetNames.Contains("0") = True Then
                        GetNames.Remove("0")
                    End If
                End If
                GetNames.Sort()
                Return GetNames
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetNames, motivo: " & ex.Message)
            End Try
        End Function


        ''' <summary>
        ''' Retorna a coleção de layers
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        '''<param name="IncludeLayerZero">Define se o layer zero será incluído</param>
        '''<param name="Filters">Filtros para seleção de camadas</param>
        ''' <param name="IsFrozen">Determina se camadas congeladas serão aceitas</param>
        ''' <param name="IsLocked">Determina se camadas bloqueadas serão aceitas</param>
        '''<returns>Boolean</returns>
        '''<remarks></remarks>
        Public Shared Function GetNames(Transaction As Transaction, Optional IncludeLayerZero As Boolean = False, Optional Filters As ArrayList = Nothing, Optional IsLocked As Object = Nothing, Optional IsFrozen As Object = Nothing) As ArrayList
            Try
                GetNames = New ArrayList
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim LayerTableRecord As LayerTableRecord
                Dim LayerTable As LayerTable
                LayerTable = Transaction.GetObject(Document.Database.LayerTableId, OpenMode.ForRead)
                For Each ObjectID As ObjectId In LayerTable
                    LayerTableRecord = DirectCast(Transaction.GetObject(ObjectID, OpenMode.ForWrite), LayerTableRecord)
                    If IsNothing(Filters) = False Then
                        For Each Filter As String In Filters
                            If LayerTableRecord.Name Like Filter = True Then
                                If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                    GetNames.Add(LayerTableRecord.Name)
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                    If LayerTableRecord.IsLocked = IsLocked Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                End If
                            End If
                        Next
                    Else
                        If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                            GetNames.Add(LayerTableRecord.Name)
                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                            If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                GetNames.Add(LayerTableRecord.Name)
                            End If
                        ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                            If LayerTableRecord.IsFrozen = IsFrozen Then
                                GetNames.Add(LayerTableRecord.Name)
                            End If
                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                            If LayerTableRecord.IsLocked = IsLocked Then
                                GetNames.Add(LayerTableRecord.Name)
                            End If
                        End If
                    End If
                Next
                If IncludeLayerZero = False Then
                    If GetNames.Contains("0") = True Then
                        GetNames.Remove("0")
                    End If
                End If
                GetNames.Sort()
                Return GetNames
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetNames, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retorna a coleção de layers
        ''' </summary>
        ''' <param name="FileName">Nome e caminho do arquivo DWG</param>
        '''<param name="IncludeLayerZero">Define se o layer zero será incluído</param>
        '''<param name="Filters">Filtros para seleção de camadas</param>
        ''' <param name="IsFrozen">Determina se camadas congeladas serão aceitas</param>
        ''' <param name="IsLocked">Determina se camadas bloqueadas serão aceitas</param>
        '''<returns>Boolean</returns>
        '''<remarks></remarks>
        Public Shared Function GetNames(FileName As Object, Optional IncludeLayerZero As Boolean = False, Optional Filters As ArrayList = Nothing, Optional IsLocked As Object = Nothing, Optional IsFrozen As Object = Nothing) As ArrayList
            Try
                GetNames = New ArrayList
                Dim LayerTableRecord As LayerTableRecord
                Dim LayerTable As LayerTable
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Using Database As New Database(False, False)
                    Database.ReadDwgFile(FileName, FileOpenMode.OpenForReadAndReadShare, False, "")
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        LayerTable = Transaction.GetObject(Database.LayerTableId, OpenMode.ForRead)
                        For Each ObjectID As ObjectId In LayerTable
                            LayerTableRecord = DirectCast(Transaction.GetObject(ObjectID, OpenMode.ForWrite), LayerTableRecord)
                            If IsNothing(Filters) = False Then
                                For Each Filter As String In Filters
                                    If LayerTableRecord.Name Like Filter = True Then
                                        If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                            GetNames.Add(LayerTableRecord.Name)
                                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                            If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                            If LayerTableRecord.IsFrozen = IsFrozen Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                            If LayerTableRecord.IsLocked = IsLocked Then
                                                GetNames.Add(LayerTableRecord.Name)
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                    GetNames.Add(LayerTableRecord.Name)
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                    If LayerTableRecord.IsFrozen = IsFrozen Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                    If LayerTableRecord.IsLocked = IsLocked Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    End If
                                End If
                            End If
                        Next
                    End Using
                End Using
                If IncludeLayerZero = False Then
                    If GetNames.Contains("0") = True Then
                        GetNames.Remove("0")
                    End If
                End If
                GetNames.Sort()
                Return GetNames
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetNames, motivo: " & ex.Message)
            End Try
        End Function


        ''' <summary>
        ''' Retorna a coleção de layers de um arquivo externo
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="FileName">Nome e caminho do arquivo DWG</param>
        ''' <param name="IncludeLayerZero">Define se o layer zero será incluído</param>
        ''' <param name="Filters">Filtros para seleção de camadas</param>
        ''' <param name="IsFrozen">Determina se camadas congeladas serão aceitas</param>
        ''' <param name="IsLocked">Determina se camadas bloqueadas serão aceitas</param>
        '''<returns>Boolean</returns>
        '''<remarks></remarks>
        Public Shared Function GetNames(Transaction As Transaction, FileName As Object, Optional IncludeLayerZero As Boolean = False, Optional Filters As ArrayList = Nothing, Optional IsLocked As Object = Nothing, Optional IsFrozen As Object = Nothing) As ArrayList
            Try
                GetNames = New ArrayList
                Dim LayerTableRecord As LayerTableRecord
                Dim LayerTable As LayerTable
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Using Database As New Database(False, False)
                    Database.ReadDwgFile(FileName, FileOpenMode.OpenForReadAndReadShare, False, "")
                    LayerTable = Transaction.GetObject(Database.LayerTableId, OpenMode.ForRead)
                    For Each ObjectID As ObjectId In LayerTable
                        LayerTableRecord = DirectCast(Transaction.GetObject(ObjectID, OpenMode.ForWrite), LayerTableRecord)
                        If IsNothing(Filters) = False Then
                            For Each Filter As String In Filters
                                If LayerTableRecord.Name Like Filter = True Then
                                    If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                        GetNames.Add(LayerTableRecord.Name)
                                    ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                        If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                            GetNames.Add(LayerTableRecord.Name)
                                        End If
                                    ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                        If LayerTableRecord.IsFrozen = IsFrozen Then
                                            GetNames.Add(LayerTableRecord.Name)
                                        End If
                                    ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                        If LayerTableRecord.IsLocked = IsLocked Then
                                            GetNames.Add(LayerTableRecord.Name)
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            If IsNothing(IsLocked) = True And IsNothing(IsFrozen) = True Then
                                GetNames.Add(LayerTableRecord.Name)
                            ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = False Then
                                If LayerTableRecord.IsLocked = IsLocked And LayerTableRecord.IsFrozen = IsFrozen Then
                                    GetNames.Add(LayerTableRecord.Name)
                                End If
                            ElseIf IsNothing(IsLocked) = True And IsNothing(IsFrozen) = False Then
                                If LayerTableRecord.IsFrozen = IsFrozen Then
                                    GetNames.Add(LayerTableRecord.Name)
                                End If
                            ElseIf IsNothing(IsLocked) = False And IsNothing(IsFrozen) = True Then
                                If LayerTableRecord.IsLocked = IsLocked Then
                                    GetNames.Add(LayerTableRecord.Name)
                                End If
                            End If
                        End If
                    Next
                End Using
                If IncludeLayerZero = False Then
                    If GetNames.Contains("0") = True Then
                        GetNames.Remove("0")
                    End If
                End If
                GetNames.Sort()
                Return GetNames
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Layer.GetNames, motivo: " & ex.Message)
            End Try
        End Function


        ''' <summary>
        ''' Limpa blocos fora de uso
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Purge() As Boolean
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim LayerTable As LayerTable
            Dim ObjectIdCollection As ObjectIdCollection
            Dim SymbolTableRecord As SymbolTableRecord
            Engine2.Layer.SetCurrent("0")
            Using DocumentLock As DocumentLock = Document.LockDocument()
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    LayerTable = Transaction.GetObject(Database.LayerTableId, OpenMode.ForRead)
                    ObjectIdCollection = New ObjectIdCollection()
                    For Each ObjectId As ObjectId In LayerTable
                        ObjectIdCollection.Add(ObjectId)
                    Next
                    Database.Purge(ObjectIdCollection)
                    For Each ObjectId As ObjectId In ObjectIdCollection
                        SymbolTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                        Try
                            SymbolTableRecord.Erase(True)
                            Return True
                        Catch
                            Return False
                        End Try
                    Next
                    Transaction.Commit()
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Limpa blocos fora de uso
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Purge(Transaction As Transaction) As Boolean
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim LayerTable As LayerTable
            Dim ObjectIdCollection As ObjectIdCollection
            Dim SymbolTableRecord As SymbolTableRecord
            Engine2.Layer.SetCurrent("0")
            LayerTable = Transaction.GetObject(Database.LayerTableId, OpenMode.ForRead)
            ObjectIdCollection = New ObjectIdCollection()
            For Each ObjectId As ObjectId In LayerTable
                ObjectIdCollection.Add(ObjectId)
            Next
            Database.Purge(ObjectIdCollection)
            For Each ObjectId As ObjectId In ObjectIdCollection
                SymbolTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                Try
                    SymbolTableRecord.Erase(True)
                    Return True
                Catch
                    Return False
                End Try
            Next
        End Function

    End Class

End Namespace


