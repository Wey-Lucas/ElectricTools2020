Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text

'Referências em http://forums.autodesk.com/t5/net/plotsettingsvalidator-setcurrentstylesheet-isn-t-affecting-line/td-p/2648006

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Gerenciamento de impressão
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AcadPrintManager

        ''' <summary>
        ''' Orientação da folha
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eOrientation
            Portrait = 0
            Landscape = 1
        End Enum

        ''' <summary>
        ''' Retorna o nome do layout corrente
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentLayout() As Object
            Try
                Return LayoutManager.Current.CurrentLayout
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Determina se a impressora plota em arquivo 
        ''' </summary>
        ''' <param name="PlotDeviceName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PlotToFile(PlotDeviceName As String) As Boolean
            If PlotDeviceName.Contains("DWF") Then
                Return True
            ElseIf PlotDeviceName.Contains("PDF") Then
                Return True
            ElseIf PlotDeviceName.Contains("JPG") Then
                Return True
            ElseIf PlotDeviceName.Contains("PNG") Then
                Return True
            ElseIf PlotDeviceName.Contains("XPS") Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Imprimir
        ''' </summary>
        ''' <param name="LayoutName">Nome do layout</param>
        ''' <param name="PlotDeviceName">Nome da impressora</param>
        ''' <param name="MediaName">Nome do papel</param>
        ''' <param name="StyleSheet">Estilo de plotagem</param>
        ''' <param name="Scale">Escala</param>
        ''' <param name="PlotType">Tipo de plotagem</param>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <param name="PlotPaperUnit">Unidade</param>
        ''' <param name="Orientation">Orientação da folha</param>
        ''' <param name="PlotUpsideDown">Define se irá imprimir de ponta cabeça</param>
        ''' <param name="Plotlineweights">Define se irá plotar as espessuras</param>
        ''' <param name="PlotSettingsShadePlotType">Define a renderização de saída</param>
        ''' <param name="PLTFile">Define se a plotagem é em PLT (Limitado a impressoras 'PlotToPrinter')</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Print(LayoutName As String, PlotDeviceName As String, MediaName As String, Optional StyleSheet As String = "acad.ctb", Optional Scale As Object = "_Fit", Optional PlotType As Autodesk.AutoCAD.DatabaseServices.PlotType = Autodesk.AutoCAD.DatabaseServices.PlotType.Window, Optional P1 As Point2d = Nothing, Optional P2 As Point2d = Nothing, Optional PlotPaperUnit As Autodesk.AutoCAD.DatabaseServices.PlotPaperUnit = PlotPaperUnit.Millimeters, Optional Orientation As eOrientation = eOrientation.Portrait, Optional PlotUpsideDown As Boolean = False, Optional Plotlineweights As Boolean = True, Optional PlotSettingsShadePlotType As Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType = PlotSettingsShadePlotType.AsDisplayed, Optional PLTFile As Boolean = False) As Boolean
            Dim PLQUIET As Integer = Application.GetSystemVariable("PLQUIET")
            Dim FILEDIA As Integer = Application.GetSystemVariable("FILEDIA")
            Dim CMDDIA As Integer = Application.GetSystemVariable("CMDDIA")
            Dim CMDECHO As Integer = Application.GetSystemVariable("CMDECHO")
            Try
                Dim SB As New StringBuilder
                Dim Path As String = ""
                Dim Extension As String = ""
                If PlotDeviceName.Contains("DWF") Then
                    Extension = ".dwf"
                ElseIf PlotDeviceName.Contains("PDF") Then
                    Extension = ".pdf"
                ElseIf PlotDeviceName.Contains("JPG") Then
                    Extension = ".jpg"
                ElseIf PlotDeviceName.Contains("PNG") Then
                    Extension = ".png"
                ElseIf PlotDeviceName.Contains("XPS") Then
                    Extension = ".xps"
                Else
                    If PLTFile = True Then
                        Extension = ".plt"
                    Else
                        Extension = ""
                    End If
                End If
                If Extension.Trim <> "" Then
                    Path = Engine2.CadDialogs.SaveFile("Destino do arquivo (*" & Extension & ")", "Arquivo *" & Extension & "|*" & Extension, "C:\", "*" & Extension)
                    If IsNothing(Path) = True Then
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                        Return False
                    Else
                        Path = Path.Replace("\", "/")
                    End If
                End If
                With SB
                    .Append("(command")
                    .Append(" ")
                    .Append(Chr(34) & "_.-PLOT" & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & "_Yes" & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & LayoutName & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & PlotDeviceName & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & MediaName & Chr(34))
                    .Append(" ")
                    Select Case PlotPaperUnit
                        Case Autodesk.AutoCAD.DatabaseServices.PlotPaperUnit.Inches
                            .Append(Chr(34) & "_Inches" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotPaperUnit.Millimeters
                            .Append(Chr(34) & "_Millimeters" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotPaperUnit.Pixels
                            .Append(Chr(34) & "_Pixels" & Chr(34))
                    End Select
                    .Append(" ")
                    Select Case Orientation
                        Case eOrientation.Landscape
                            .Append(Chr(34) & "_Landscape" & Chr(34))
                        Case eOrientation.Portrait
                            .Append(Chr(34) & "_Portrait" & Chr(34))
                    End Select
                    .Append(" ")
                    If PlotUpsideDown = False Then
                        .Append(Chr(34) & "_No" & Chr(34))
                    Else
                        .Append(Chr(34) & "_Yes" & Chr(34))
                    End If
                    .Append(" ")
                    Select Case PlotType
                        Case Autodesk.AutoCAD.DatabaseServices.PlotType.Display
                            .Append(Chr(34) & "_Display" & Chr(34))
                            .Append(" ")
                        Case Autodesk.AutoCAD.DatabaseServices.PlotType.Extents
                            .Append(Chr(34) & "_Extents" & Chr(34))
                            .Append(" ")
                        Case Autodesk.AutoCAD.DatabaseServices.PlotType.Limits
                            .Append(Chr(34) & "_Limits" & Chr(34))
                            .Append(" ")
                        Case Autodesk.AutoCAD.DatabaseServices.PlotType.View
                            .Append(Chr(34) & "_View" & Chr(34))
                            .Append(" ")
                        Case Autodesk.AutoCAD.DatabaseServices.PlotType.Window
                            .Append(Chr(34) & "_Window" & Chr(34))
                            .Append(" ")
                            .Append("(List " & P1.X & " " & P1.Y & ")")
                            .Append(" ")
                            .Append("(List " & P2.X & " " & P2.Y & ")")
                            .Append(" ")
                        Case Else
                            Throw New System.Exception("PlotType = " & PlotType.ToString & " não emplementado")
                    End Select
                    .Append(Chr(34) & Scale & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & "_Center" & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & "_Yes" & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & StyleSheet & Chr(34))
                    .Append(" ")
                    If Plotlineweights = True Then
                        .Append(Chr(34) & "_Yes" & Chr(34))
                    Else
                        .Append(Chr(34) & "_No" & Chr(34))
                    End If
                    .Append(" ")
                    Select Case PlotSettingsShadePlotType
                        Case Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType.AsDisplayed
                            .Append(Chr(34) & "_A" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType.Hidden
                            .Append(Chr(34) & "_H" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType.Rendered
                            .Append(Chr(34) & "_R" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType.VisualStyle
                            .Append(Chr(34) & "_V" & Chr(34))
                        Case Autodesk.AutoCAD.DatabaseServices.PlotSettingsShadePlotType.Wireframe
                            .Append(Chr(34) & "_W" & Chr(34))
                        Case Else
                            Throw New System.Exception("PlotSettingsShadePlotType = " & PlotSettingsShadePlotType.ToString & " não emplementado")
                    End Select
                    .Append(" ")
                    If Extension.Trim <> "" Then
                        .Append(Chr(34) & Path & Chr(34))
                    Else
                        If PLTFile = True Then
                            .Append(Chr(34) & "_Yes" & Chr(34))
                            .Append(" ")
                            .Append(Chr(34) & Path & Chr(34))
                        Else
                            .Append(Chr(34) & "_No" & Chr(34))
                        End If
                    End If
                    .Append(" ")
                    .Append(Chr(34) & "_No" & Chr(34))
                    .Append(" ")
                    .Append(Chr(34) & "_Yes" & Chr(34))
                    .Append(")")
                    .Append("(princ) ")
                End With
                Application.SetSystemVariable("PLQUIET", 1)
                Application.SetSystemVariable("FILEDIA", 0)
                Application.SetSystemVariable("CMDDIA", 0)
                Application.SetSystemVariable("CMDECHO", 0)
                Application.DocumentManager.MdiActiveDocument.SendStringToExecute(SB.ToString, True, False, False)
                Return True
            Catch
                Return False
            Finally
                Application.SetSystemVariable("PLQUIET", PLQUIET)
                Application.SetSystemVariable("FILEDIA", FILEDIA)
                Application.SetSystemVariable("CMDDIA", CMDDIA)
                Application.SetSystemVariable("CMDECHO", CMDECHO)
            End Try
        End Function

        ''' <summary>
        ''' Impressoras
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PlotDeviceNames() As ArrayList
            Try
                PlotDeviceNames = New ArrayList
                Dim AcadApplication As AcadApplication = TryCast(Application.AcadApplication, AcadApplication)
                Dim AcadLayout As Common.AcadLayout = AcadApplication.ActiveDocument.ActiveLayout
                For Each DeviceName As String In AcadLayout.GetPlotDeviceNames()
                    PlotDeviceNames.Add(DeviceName)
                Next
                PlotDeviceNames.Sort()
                Return PlotDeviceNames
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna a lista de arquivos *.ctb
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SheetStyles() As ArrayList
            Try
                SheetStyles = New ArrayList
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim layManager As LayoutManager
                Dim ObjectId As ObjectId
                Dim Layout As Layout
                Dim PlotSettingsValidator As PlotSettingsValidator
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    layManager = LayoutManager.Current
                    ObjectId = layManager.GetLayoutId(layManager.CurrentLayout)
                    Layout = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                    PlotSettingsValidator = PlotSettingsValidator.Current
                    PlotSettingsValidator.RefreshLists(Layout)
                    For Each SheetName As String In PlotSettingsValidator.GetPlotStyleSheetList()
                        SheetStyles.Add(SheetName)
                    Next
                End Using
                SheetStyles.Sort()
                Return SheetStyles
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna as folhas configuradas para impressora
        ''' </summary>
        ''' <param name="PlotDeviceName">Nome da impressora</param>
        ''' <param name="LocalMediaName">Determina o formato de saída para os nomes de folha, Use LocalMediaName = True para preenchimento de coleções ou LocalMediaName = False para uso de AcadPrintManager.Print</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MediaNames(PlotDeviceName As String, Optional LocalMediaName As Boolean = False) As ArrayList
            Try
                MediaNames = New ArrayList
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Item As Integer = 0
                Dim PlotSettingsValidator As PlotSettingsValidator
                Using PlotSettings As PlotSettings = New PlotSettings(True)
                    PlotSettingsValidator = PlotSettingsValidator.Current
                    PlotSettingsValidator.SetPlotConfigurationName(PlotSettings, PlotDeviceName, Nothing)
                    For Each MediaName As String In PlotSettingsValidator.GetCanonicalMediaNameList(PlotSettings)
                        If LocalMediaName = True Then
                            MediaNames.Add(PlotSettingsValidator.GetLocaleMediaName(PlotSettings, Item))
                        Else
                            MediaNames.Add(MediaName)
                        End If
                        Item = Item + 1
                    Next
                End Using
                Return MediaNames
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna o nome local de uma folha
        ''' </summary>
        ''' <param name="PlotDeviceName">Nome da impressora</param>
        ''' <param name="MediaName">Nome da folha (Item de coleção retornado pela função MediaNames onde o parâmetro LocalMediaName = False)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LocalMediaName(PlotDeviceName As String, MediaName As String) As Object
            Try
                LocalMediaName = Nothing
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Item As Integer = 0
                Dim PlotSettingsValidator As PlotSettingsValidator
                Using PlotSettings As PlotSettings = New PlotSettings(True)
                    PlotSettingsValidator = PlotSettingsValidator.Current
                    PlotSettingsValidator.SetPlotConfigurationName(PlotSettings, PlotDeviceName, Nothing)
                    LocalMediaName = PlotSettingsValidator.GetLocaleMediaName(PlotSettings, MediaName)
                End Using
                Return LocalMediaName
            Catch
                Return Nothing
            End Try
        End Function





























        '    '    Imports Autodesk.AutoCAD.Runtime
        '    'Imports Autodesk.AutoCAD.ApplicationServices
        '    'Imports Autodesk.AutoCAD.DatabaseServices
        '    'Imports Autodesk.AutoCAD.EditorInput
        '    'Imports Autodesk.AutoCAD.PlottingServices
        '    'Imports Autodesk.AutoCAD.Geometry
        '    'Imports System.Runtime.InteropServices


        '    Public Class SimplePlottingCommands
        '        <DllImport("acad.exe", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedTrans")> _
        '        Private Shared Function acedTrans(point As Double(), fromRb As IntPtr, toRb As IntPtr, disp As Integer, result As Double()) As Integer
        '        End Function

        '        <CommandMethod("winplot")> _
        '        Public Shared Sub WindowPlot()
        '            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        '            Dim ed As Editor = doc.Editor
        '            Dim db As Database = doc.Database

        '            Dim ppo As New PromptPointOptions(vbLf & "Select first corner of plot area: ")
        '            ppo.AllowNone = False

        '            Dim ppr As PromptPointResult = ed.GetPoint(ppo)

        '            If ppr.Status <> PromptStatus.OK Then
        '                Return
        '            End If

        '            Dim first As Point3d = ppr.Value

        '            Dim pco As New PromptCornerOptions(vbLf & "Select second corner of plot area: ", first)
        '            ppr = ed.GetCorner(pco)

        '            If ppr.Status <> PromptStatus.OK Then
        '                Return
        '            End If

        '            Dim second As Point3d = ppr.Value

        '            ' Transform from UCS to DCS

        '            Dim rbFrom As New ResultBuffer(New TypedValue(5003, 1)), rbTo As New ResultBuffer(New TypedValue(5003, 2))

        '            Dim firres As Double() = New Double() {0, 0, 0}
        '            Dim secres As Double() = New Double() {0, 0, 0}

        '            ' Transform the first point...

        '            acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres)

        '            ' ... and the second

        '            acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres)

        '            ' We can safely drop the Z-coord at this stage

        '            Dim window As New Extents2d(firres(0), firres(1), secres(0), secres(1))

        '            Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '            Using tr
        '                ' We'll be plotting the current layout

        '                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead), BlockTableRecord)
        '                Dim lo As Layout = DirectCast(tr.GetObject(btr.LayoutId, OpenMode.ForRead), Layout)

        '                ' We need a PlotInfo object
        '                ' linked to the layout

        '                Dim pi As New PlotInfo()
        '                pi.Layout = btr.LayoutId

        '                ' We need a PlotSettings object
        '                ' based on the layout settings
        '                ' which we then customize

        '                Dim ps As New PlotSettings(lo.ModelType)
        '                ps.CopyFrom(lo)

        '                ' The PlotSettingsValidator helps
        '                ' create a valid PlotSettings object

        '                Dim psv As PlotSettingsValidator = PlotSettingsValidator.Current

        '                ' We'll plot the extents, centered and
        '                ' scaled to fit

        '                psv.SetPlotWindowArea(ps, window)
        '                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window)
        '                psv.SetUseStandardScale(ps, True)
        '                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit)
        '                psv.SetPlotCentered(ps, True)

        '                ' We'll use the standard DWF PC3, as
        '                ' for today we're just plotting to file

        '                psv.SetPlotConfigurationName(ps, "DWF6 ePlot.pc3", "ANSI_A_(8.50_x_11.00_Inches)")

        '                ' We need to link the PlotInfo to the
        '                ' PlotSettings and then validate it

        '                pi.OverrideSettings = ps
        '                Dim piv As New PlotInfoValidator()
        '                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled
        '                piv.Validate(pi)

        '                ' A PlotEngine does the actual plotting
        '                ' (can also create one for Preview)

        '                If PlotFactory.ProcessPlotState = ProcessPlotState.NotPlotting Then
        '                    Dim pe As PlotEngine = PlotFactory.CreatePublishEngine()
        '                    Using pe
        '                        ' Create a Progress Dialog to provide info
        '                        ' and allow thej user to cancel

        '                        Dim ppd As New PlotProgressDialog(False, 1, True)
        '                        Using ppd
        '                            ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress")
        '                            ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job")
        '                            ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet")
        '                            ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress")
        '                            ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress")
        '                            ppd.LowerPlotProgressRange = 0
        '                            ppd.UpperPlotProgressRange = 100
        '                            ppd.PlotProgressPos = 0

        '                            ' Let's start the plot, at last

        '                            ppd.OnBeginPlot()
        '                            ppd.IsVisible = True
        '                            pe.BeginPlot(ppd, Nothing)

        '                            ' We'll be plotting a single document

        '                            ' Let's plot to file
        '                            pe.BeginDocument(pi, doc.Name, Nothing, 1, True, "c:\test-output")

        '                            ' Which contains a single sheet

        '                            ppd.OnBeginSheet()

        '                            ppd.LowerSheetProgressRange = 0
        '                            ppd.UpperSheetProgressRange = 100
        '                            ppd.SheetProgressPos = 0

        '                            Dim ppi As New PlotPageInfo()
        '                            pe.BeginPage(ppi, pi, True, Nothing)
        '                            pe.BeginGenerateGraphics(Nothing)
        '                            pe.EndGenerateGraphics(Nothing)

        '                            ' Finish the sheet
        '                            pe.EndPage(Nothing)
        '                            ppd.SheetProgressPos = 100
        '                            ppd.OnEndSheet()

        '                            ' Finish the document

        '                            pe.EndDocument(Nothing)

        '                            ' And finish the plot

        '                            ppd.PlotProgressPos = 100
        '                            ppd.OnEndPlot()
        '                            pe.EndPlot(Nothing)
        '                        End Using
        '                    End Using
        '                Else
        '                    ed.WriteMessage(vbLf & "Another plot is in progress.")
        '                End If
        '            End Using
        '        End Sub
        '    End Class

    End Class

End Namespace
