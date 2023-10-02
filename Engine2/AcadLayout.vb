Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Collections.Specialized
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text
Imports System.Text.RegularExpressions

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2018
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    Public Class AcadLayout

        'Editor.SwitchToModelSpace()
        'Editor.SwitchToModelSpace()

        ''' <summary>
        ''' Retorna a viewport ativa
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetActiveSpace() As String
            'Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            'Dim Editor As Editor = Document.Editor
            'Dim Database As Database = Document.Database
            'Dim SpaceId As ObjectId = Database.CurrentSpaceId
            'Dim ModelSpaceId As ObjectId = SymbolUtilityServices.GetBlockModelSpaceId(Database)
            Dim layoutMgr As LayoutManager = LayoutManager.Current
            Return layoutMgr.CurrentLayout
        End Function

        ''' <summary>
        ''' Exibe todos os layouts
        ''' </summary>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListLayouts(Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As List(Of Layout)
            Dim Database As Database
            Dim Editor As Editor
            Dim DBDictionary As DBDictionary
            Dim Layout As Layout
            ListLayouts = New List(Of Layout)
            Try
                If IsNothing(Document) = True Then
                    Document = Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
                Database = Document.Database
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    DBDictionary = Transaction.GetObject(Database.LayoutDictionaryId, OpenMode.ForRead)
                    For Each Dict As DBDictionaryEntry In DBDictionary
                        Layout = Transaction.GetObject(CType(DBDictionary(Dict.Key), ObjectId), OpenMode.ForRead)
                        ListLayouts.Add(Layout)
                    Next
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction()
                        Using Transaction
                            ListLayouts = ListLayouts(Document, Transaction, ActiveDocument)
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return ListLayouts
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Exclui todos os layouts
        ''' </summary>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <param name="Exceptions">Relação de nomes de layouts a serem preservados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseAllLayouts(Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True, Optional Exceptions As List(Of String) = Nothing) As Boolean
            Dim Database As Database
            Dim Editor As Editor
            Dim DBDictionary As DBDictionary
            Try
                If IsNothing(Document) = True Then
                    Document = Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
                Database = Document.Database
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    If IsNothing(Exceptions) = True Then
                        Exceptions = New List(Of String)
                    End If
                    If Exceptions.Contains("Model") = False Then
                        Exceptions.Add("Model")
                    End If
                    DBDictionary = Transaction.GetObject(Database.LayoutDictionaryId, OpenMode.ForRead)
                    For Each Dict As DBDictionaryEntry In DBDictionary
                        If Exceptions.Contains(Dict.Key) = False Then
                            LayoutManager.Current.DeleteLayout(Dict.Key)
                        End If
                    Next
                    Editor.Regen()
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction()
                        Using Transaction
                            EraseAllLayouts(Document, Transaction, ActiveDocument, Exceptions)
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Cria um novo layout
        ''' </summary>
        ''' <param name="LayoutName">Nome do novo layout</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <param name="DrawViewportsFirst">Determina se a viewport será criada no layout</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateLayout(LayoutName As String, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True, Optional DrawViewportsFirst As Boolean = False) As Layout
            Dim Database As Database
            Dim Editor As Editor
            Dim LayoutManager As LayoutManager
            Dim ObjectId As ObjectId
            Dim Layout As Layout
            Dim Viewport As Viewport
            Try
                Layout = Nothing
                If IsNothing(Document) = True Then
                    Document = Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
                Database = Document.Database
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    LayoutManager = LayoutManager.Current
                    ObjectId = LayoutManager.CreateLayout(LayoutName)
                    Layout = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                    LayoutManager.CurrentLayout = Layout.LayoutName
                    If DrawViewportsFirst = False Then
                        Layout.UpgradeOpen()
                        Layout.DrawViewportsFirst = DrawViewportsFirst
                        For Each ID As ObjectId In Layout.GetViewports
                            Viewport = Transaction.GetObject(ID, OpenMode.ForWrite)
                            Viewport.Erase()
                        Next
                    End If
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction()
                        Using Transaction
                            Layout = CreateLayout(LayoutName, Document, Transaction)
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return Layout
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem a máxima extensão para o layout
        ''' </summary>
        ''' <param name="Layout">Layout</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetMaximumExtents(ByVal Layout As Layout) As Extents2d
            Dim Min As Point2d
            Dim Max As Point2d
            Dim Units As Double
            Try
                Select Case Layout.PlotPaperUnits
                    Case PlotPaperUnit.Millimeters
                        Units = 25.4
                    Case PlotPaperUnit.Millimeters, PlotPaperUnit.Pixels
                        Units = 1
                End Select
                Select Case Layout.PlotRotation
                    Case PlotRotation.Degrees090, PlotRotation.Degrees270
                        Min = New Point2d(Layout.PlotPaperMargins.MinPoint.Y, Layout.PlotPaperMargins.MinPoint.X) / Units
                        Max = (New Point2d(Layout.PlotPaperSize.Y, Layout.PlotPaperSize.X) - New Point2d(Layout.PlotPaperMargins.MaxPoint.Y, Layout.PlotPaperMargins.MaxPoint.X).GetAsVector) / Units
                    Case Else
                        Min = Layout.PlotPaperMargins.MinPoint / Units
                        Max = Layout.PlotPaperSize / Units
                End Select
                Return New Extents2d(Min, Max)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Seta os dados de plotagem
        ''' </summary>
        ''' <param name="PageSetupName"></param>
        ''' <param name="Layout"></param>
        ''' <param name="PageSize"></param>
        ''' <param name="StyleSheet"></param>
        ''' <param name="Device"></param>
        ''' <param name="Offset"></param>
        ''' <param name="CustomScale"></param>
        ''' <param name="PlotPaperUnit"></param>
        ''' <param name="SetUseStandardScale"></param>
        ''' <param name="StdScaleType"></param>
        ''' <param name="PlotRotation"></param>
        ''' <param name="ShadePlotResLevel"></param>
        ''' <param name="ShadePlot"></param>
        ''' <param name="ScaleLineweights"></param>
        ''' <param name="ShowPlotStyles"></param>
        ''' <param name="PrintLineweights"></param>
        ''' <param name="PlotTransparency"></param>
        ''' <param name="PlotPlotStyles"></param>
        ''' <param name="DrawViewportsFirst"></param>
        ''' <param name="PlotHidden"></param>
        ''' <param name="PlotType"></param>
        ''' <param name="PlotArea"></param>
        ''' <param name="PlotCentered"></param>
        ''' <param name="Document"></param>
        ''' <param name="Transaction"></param>
        ''' <param name="ActiveDocument"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SetPlotSettings(
                                              PageSetupName As String,
                                              Layout As Layout,
                                              PageSize As String,
                                              StyleSheet As String,
                                              Device As String,
                                              Offset As Point2d,
                                              CustomScale As CustomScale,
                                              PlotPaperUnit As PlotPaperUnit,
                                              Optional SetUseStandardScale As Boolean = False,
                                              Optional StdScaleType As StdScaleType = StdScaleType.ScaleToFit,
                                              Optional PlotRotation As PlotRotation = PlotRotation.Degrees090,
                                              Optional ShadePlotResLevel As ShadePlotResLevel = ShadePlotResLevel.Normal,
                                              Optional ShadePlot As PlotSettingsShadePlotType = PlotSettingsShadePlotType.AsDisplayed,
                                              Optional ScaleLineweights As Boolean = False,
                                              Optional ShowPlotStyles As Boolean = False,
                                              Optional PrintLineweights As Boolean = True,
                                              Optional PlotTransparency As Boolean = False,
                                              Optional PlotPlotStyles As Boolean = True,
                                              Optional DrawViewportsFirst As Boolean = True,
                                              Optional PlotHidden As Boolean = False,
                                              Optional PlotType As PlotType = PlotType.Extents,
                                              Optional PlotArea As Extents2d = Nothing,
                                              Optional PlotCentered As Boolean = False,
                                              Optional Document As Document = Nothing,
                                              Optional Transaction As Transaction = Nothing,
                                              Optional ActiveDocument As Boolean = True
                                                                                    ) As PlotSettings
            Dim Database As Database
            Dim Editor As Editor
            Dim PlotSettingsValidator As PlotSettingsValidator
            Dim PlotDevices As StringCollection
            Dim PlotPaper As StringCollection
            Dim PlotStyle As StringCollection
            Dim PlotSettings As PlotSettings
            Dim IsNew As Boolean = False
            Dim PlotSettingsDictionary As DBDictionary
            Dim LayoutManager As LayoutManager
            Try
                SetPlotSettings = Nothing
                If IsNothing(Document) = True Then
                    Document = Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        Application.DocumentManager.MdiActiveDocument = Document
                    End If
                End If
                Database = Document.Database
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    PlotSettingsDictionary = Transaction.GetObject(Database.PlotSettingsDictionaryId, OpenMode.ForRead)
                    LayoutManager = LayoutManager.Current
                    Layout = Transaction.GetObject(Layout.Id, OpenMode.ForWrite)
                    If PlotSettingsDictionary.Contains(PageSetupName) = False Then
                        IsNew = True
                        PlotSettings = New PlotSettings(Layout.ModelType)
                        PlotSettings.CopyFrom(Layout)
                        PlotSettings.PlotSettingsName = PageSetupName
                        'CRIA UM NOVO PlotSettings (OMITIMOS ISTO PARA ATUALIZAR O ATUAL)
                        'PlotSettings.AddToPlotSettingsDictionary(Database)
                        'Transaction.AddNewlyCreatedDBObject(PlotSettings, True)
                    Else
                        PlotSettings = PlotSettingsDictionary.GetAt(PageSetupName).GetObject(OpenMode.ForWrite)
                    End If
                    Using PlotSettings
                        PlotSettingsValidator = PlotSettingsValidator.Current
                        PlotDevices = PlotSettingsValidator.GetPlotDeviceList()
                        If PlotDevices.Contains(Device) Then
                            PlotSettingsValidator.SetPlotConfigurationName(PlotSettings, Device, Nothing)
                        End If
                        PlotSettingsValidator.RefreshLists(PlotSettings)
                        PageSize = GetPageSize(PlotSettings, PlotSettingsValidator, PageSize)
                        PlotPaper = PlotSettingsValidator.GetCanonicalMediaNameList(PlotSettings)
                        If IsNothing(PageSize) = False Then
                            If PlotPaper.Contains(PageSize) Then
                                PlotSettingsValidator.SetCanonicalMediaName(PlotSettings, PageSize)
                            End If
                        End If
                        PlotStyle = PlotSettingsValidator.GetPlotStyleSheetList()
                        If PlotStyle.Contains(StyleSheet) Then
                            PlotSettingsValidator.SetCurrentStyleSheet(PlotSettings, StyleSheet)
                        End If
                        PlotSettingsValidator.SetCustomPrintScale(PlotSettings, CustomScale)
                        PlotSettingsValidator.SetPlotOrigin(PlotSettings, Offset)
                        If Layout.ModelType = True Then
                            Select Case PlotType
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Display
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Limits
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Window
                                    PlotSettingsValidator.SetPlotWindowArea(PlotSettings, PlotArea)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                            End Select
                        Else
                            Select Case PlotType
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Display
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Extents
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Layout
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, False)
                                Case Autodesk.AutoCAD.DatabaseServices.PlotType.Window
                                    PlotSettingsValidator.SetPlotWindowArea(PlotSettings, PlotArea)
                                    PlotSettingsValidator.SetPlotType(PlotSettings, PlotType)
                                    PlotSettingsValidator.SetPlotCentered(PlotSettings, PlotCentered)
                            End Select
                        End If
                        If SetUseStandardScale = True Then
                            PlotSettingsValidator.SetUseStandardScale(PlotSettings, SetUseStandardScale)
                            PlotSettingsValidator.SetStdScaleType(PlotSettings, StdScaleType)
                        End If
                        PlotSettings.ScaleLineweights = ScaleLineweights
                        PlotSettings.ShowPlotStyles = ShowPlotStyles
                        PlotSettings.ShadePlot = ShadePlot
                        PlotSettings.ShadePlotResLevel = ShadePlotResLevel
                        PlotSettings.PrintLineweights = PrintLineweights
                        PlotSettings.PlotTransparency = PlotTransparency
                        PlotSettings.PlotPlotStyles = PlotPlotStyles
                        PlotSettings.DrawViewportsFirst = DrawViewportsFirst
                        PlotSettings.PlotHidden = PlotHidden
                        PlotSettingsValidator.SetPlotRotation(PlotSettings, PlotRotation)
                        PlotSettingsValidator.RefreshLists(PlotSettings)
                        PlotSettingsValidator.SetPlotPaperUnits(PlotSettings, PlotPaperUnit)
                        Layout.CopyFrom(PlotSettings)
                        PlotSettingsValidator.SetZoomToPaperOnUpdate(PlotSettings, True)
                    End Using
                    SetPlotSettings = PlotSettings
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction()
                        Using Transaction
                            SetPlotSettings = SetPlotSettings(
                                PageSetupName,
                                Layout,
                                PageSize,
                                StyleSheet,
                                Device,
                                Offset,
                                CustomScale,
                                PlotPaperUnit,
                                SetUseStandardScale,
                                StdScaleType,
                                PlotRotation,
                                ShadePlotResLevel,
                                ShadePlot,
                                ScaleLineweights,
                                ShowPlotStyles,
                                PrintLineweights,
                                PlotTransparency,
                                PlotPlotStyles,
                                DrawViewportsFirst,
                                PlotHidden,
                                PlotType,
                                PlotArea,
                                PlotCentered,
                                Document,
                                Transaction,
                                ActiveDocument
                                )
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return SetPlotSettings
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem as dimensões da folha
        ''' </summary>
        ''' <param name="PageSetupName"></param>
        ''' <param name="Layout"></param>
        ''' <param name="Device"></param>
        ''' <param name="PageSize"></param>
        ''' <param name="PlotRotation"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPageDimension(PageSetupName As String, Layout As Layout, Device As String, PageSize As String, Optional PlotRotation As PlotRotation = PlotRotation.Degrees090) As System.Drawing.SizeF
            Dim Database As Database
            Dim Editor As Editor
            Dim PlotSettingsValidator As PlotSettingsValidator
            Dim PlotSettings As PlotSettings
            Dim Sizef As System.Drawing.SizeF = Nothing
            Dim PlotSettingsDictionary As DBDictionary
            Dim LayoutManager As LayoutManager
            Dim PlotDevices As StringCollection
            Dim Document As Document
            Dim Transaction As Transaction
            Dim PlotPaper As StringCollection
            Try
                Document = Application.DocumentManager.MdiActiveDocument
                Database = Document.Database
                Editor = Document.Editor
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        PlotSettingsDictionary = Transaction.GetObject(Database.PlotSettingsDictionaryId, OpenMode.ForRead)
                        LayoutManager = LayoutManager.Current
                        Layout = Transaction.GetObject(Layout.Id, OpenMode.ForRead)
                        If PlotSettingsDictionary.Contains(PageSetupName) = False Then
                            PlotSettings = New PlotSettings(Layout.ModelType)
                            PlotSettings.CopyFrom(Layout)
                        Else
                            PlotSettings = PlotSettingsDictionary.GetAt(PageSetupName).GetObject(OpenMode.ForRead)
                        End If
                        Using PlotSettings
                            PlotSettingsValidator = PlotSettingsValidator.Current
                            PlotDevices = PlotSettingsValidator.GetPlotDeviceList()
                            If PlotDevices.Contains(Device) Then
                                PlotSettingsValidator.SetPlotConfigurationName(PlotSettings, Device, Nothing)
                            End If
                            PlotSettingsValidator.RefreshLists(PlotSettings)
                            PlotSettingsValidator = PlotSettingsValidator.Current
                            PlotDevices = PlotSettingsValidator.GetPlotDeviceList()
                            If PlotDevices.Contains(Device) Then
                                PlotSettingsValidator.SetPlotConfigurationName(PlotSettings, Device, Nothing)
                            End If
                            PlotSettingsValidator.RefreshLists(PlotSettings)
                            PageSize = GetPageSize(PlotSettings, PlotSettingsValidator, PageSize)
                            PlotPaper = PlotSettingsValidator.GetCanonicalMediaNameList(PlotSettings)
                            If IsNothing(PageSize) = False Then
                                If PlotPaper.Contains(PageSize) Then
                                    PlotSettingsValidator.SetCanonicalMediaName(PlotSettings, PageSize)
                                End If
                            End If
                            Select Case PlotRotation
                                Case PlotRotation.Degrees090, PlotRotation.Degrees270
                                    Sizef = New System.Drawing.SizeF(PlotSettings.PlotPaperSize.Y, PlotSettings.PlotPaperSize.X)
                                Case PlotRotation.Degrees000, PlotRotation.Degrees180
                                    Sizef = New System.Drawing.SizeF(PlotSettings.PlotPaperSize.X, PlotSettings.PlotPaperSize.Y)
                            End Select
                        End Using
                        Transaction.Abort()
                    End Using
                End Using
                Return Sizef
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem valores entre parênteses
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetParenthesesValue(Value As String) As String
            Dim GroupCollection As GroupCollection
            Try
                If Value.Contains("(") = True And Value.Contains(")") = True Then
                    GroupCollection = Regex.Match(Value, "\([^)]*\)").Groups()
                    If GroupCollection.Count > 0 Then
                        Return GroupCollection(0).ToString.Replace("(", "").Replace(")", "").Trim
                    Else
                        Return ""
                    End If
                Else
                    Return Value
                End If
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Obtem o nome raiz do arquivo de folha customizado
        ''' </summary>
        ''' <param name="PlotSettings"></param>
        ''' <param name="PlotSettingsValidator"></param>
        ''' <param name="PageSize"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetPageSize(PlotSettings As PlotSettings, PlotSettingsValidator As PlotSettingsValidator, PageSize As String) As Object
            Dim LocaleName As String
            Dim CanonicalName As String
            Try
                With PlotSettingsValidator
                    For Index As Integer = 0 To .GetCanonicalMediaNameList(PlotSettings).Count - 1
                        CanonicalName = .GetCanonicalMediaNameList(PlotSettings)(Index)
                        LocaleName = .GetLocaleMediaName(PlotSettings, Index)
                        If CanonicalName.Equals(PageSize, StringComparison.OrdinalIgnoreCase) = True Or LocaleName.Equals(PageSize, StringComparison.OrdinalIgnoreCase) = True Then
                            Return CanonicalName
                        End If
                    Next
                End With
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace
