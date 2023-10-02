Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports ElectricTools2020.Engine2.AcadInterface.SelectSequenceItem
Imports ElectricTools2020.Engine2.EntityInteration.IntersectionEntityCollection
Imports ElectricTools2020.Engine2.EntityInteration
Imports Autodesk.AutoCAD.EditorInput.Editor
Imports Autodesk.AutoCAD.Interop
Imports ElectricTools2020.Engine2.Geometry

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe de interação com o usuário
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AcadInterface

        ''' <summary>
        ''' Regras para tratamento de blocos
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eBlockSegmentRules
            ''' <summary>
            ''' Considera apenas o ponto de inserção como ponto de ligação para outras entidades
            ''' </summary>
            ''' <remarks></remarks>
            ConsiderOnlyInsertionPointToBlocks = 0

            ''' <summary>
            ''' Considera qualquer parte do bloco como ponto de ligação para outras entidades
            ''' </summary>
            ''' <remarks></remarks>
            ConsiderAllBlock = 1
        End Enum


        ''' <summary>
        ''' Modo de seleção para SelectReach
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eSelectReachMode
            ''' <summary>
            ''' Poligono de seleção onde o que toca e o que esta contido será considerado
            ''' </summary>
            ''' <remarks></remarks>
            SelectCrossingPolygon = 0

            ''' <summary>
            ''' Segmento onde o que toca ou cruza será considerado
            ''' </summary>
            ''' <remarks></remarks>
            SelectFence = 1

            ''' <summary>
            ''' Poligono de seleção onde apenas o que esta inteiramente contido será considerado
            ''' </summary>
            ''' <remarks></remarks>
            SelectWindowPolygon = 2
        End Enum

        ''' <summary>
        ''' Retorna a lista com os nomes das classes
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetClassNames(TypedValueCollection As List(Of TypedValue)) As ArrayList
            GetClassNames = New ArrayList
            For Each TypedValue As TypedValue In TypedValueCollection
                With GetClassNames
                    If TypedValue.TypeCode = 0 Then
                        Select Case TypedValue.Value.ToString.ToUpper.Trim
                            Case "INSERT"
                                .Add("BlockReference")
                            Case "LINE"
                                .Add("Line")
                            Case "*LINE"
                                .Add("Line")
                                .Add("Mlime")
                                .Add("Polyline")
                                .Add("Polyline2d")
                                .Add("Polyline3d")
                            Case "POLYLINE"
                                .Add("Polyline")
                            Case "POLYLINE*"
                                .Add("Polyline")
                                .Add("Polyline2d")
                                .Add("Polyline3d")
                            Case "LWPOLYLINE", "*POLYLINE"
                                .Add("Polyline")
                            Case "ARC"
                                .Add("Arc")
                            Case "MLINE"
                                .Add("Mline")
                            Case "ELLIPSE"
                                .Add("Ellipse")
                            Case "CIRCLE"
                                .Add("Circle")
                            Case "MTEXT"
                                .Add("Mtext")
                            Case "DTEXT"
                                .Add("DBText")
                            Case "*TEXT"
                                .Add("Mtext")
                                .Add("DBText")
                            Case "POINT"
                                .Add("DBPoint")
                            Case "SPLINE"
                                .Add("Spline")
                            Case Else
                                Throw New System.Exception("Parâmetro '" & TypedValue.Value.ToString.ToUpper.Trim & "' não previsto para Engine.AcadInterface.GetClassNames")
                        End Select
                    End If
                End With
            Next
            Return GetClassNames
        End Function

        ''' <summary>
        ''' Cancela o comando corrente
        ''' </summary>
        ''' <param name="Force">Força o cancelamento do comando se nenhum comando for detectado</param>
        ''' <param name="CancelCount">Número de solicitações de cancelamento para o parâmetro Force, importante para subcomandos</param>
        Public Shared Sub CancelCommand(Optional Force As Boolean = False, Optional CancelCount As Integer = 2)

            'Declarações
            Dim Document As Document
            Dim Escape As String = ""
            Dim Commands As String
            Dim Count As Integer

            'Obtem o documento
            Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            'Obtem os comandos ativos
            Commands = CStr(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDNAMES"))

            'Avalia
            If Commands.Length > 0 Then

                'Obtem a quantidade de comandos
                Count = Commands.Split(New Char() {"'"c}).Length

                'Percorre a coleção
                For i As Integer = 0 To Count - 1

                    'Monta o comando de cancelamento
                    Escape += ChrW(3)

                Next

            Else

                'Avalia se vai forçar o cancelamento
                If Force = True Then

                    'Percorre a coleção
                    For i As Integer = 1 To CancelCount

                        'Monta o comando de cancelamento
                        Escape += ChrW(3)

                    Next

                End If

            End If

            'Avalia
            If Escape.Length > 0 Then

                'Executa o cancelamento
                Document.SendStringToExecute(Escape, True, False, True)

                'Exibe a mensagem
                Document.Editor.WriteMessage(vbCr & "*Cancel*" & vbCr)

            End If

        End Sub

        ''' <summary>
        ''' Retorna a última entidade
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Entlast() As Object
            Return Autodesk.AutoCAD.Internal.Utils.EntLast()
        End Function

        ''' <summary>
        ''' Retorna a primeira entidade
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EntFirst() As Object
            Return Autodesk.AutoCAD.Internal.Utils.EntFirst()
        End Function

        ''' <summary>
        ''' Retorna a próxima entidade a partir de uma entidade de referência
        ''' </summary>
        ''' <param name="ObjectId">Id da entidade de referência</param>
        ''' <param name="SkipSubEnt">Determina se o sistema irá considerar subentidades</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EntNext(ObjectId As ObjectId, Optional SkipSubEnt As Boolean = False) As Object
            Return Autodesk.AutoCAD.Internal.Utils.EntNext(ObjectId, SkipSubEnt)
        End Function

        'Public Shared Function SelectAtCurve(Curve As Curve, Optional TypedValueCollection As List(Of TypedValue) = Nothing) As Dictionary(Of ObjectId, Point3dCollection)

        '    Dim Result As New Dictionary(Of ObjectId, Point3dCollection)

        '    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim Editor As Editor = Document.Editor
        '    Dim Database As Database = Document.Database
        '    Dim PromptNestedEntityOptions As New PromptNestedEntityOptions("")
        '    Dim PromptNestedEntityResult As PromptNestedEntityResult
        '    Dim Elements As New Dictionary(Of ObjectId, List(Of Point3dCollection))
        '    Dim PromptSelectionResult As PromptSelectionResult
        '    Dim Point3dcollection As Point3dCollection
        '    Dim Point3d As Point3d
        '    Dim Lenght As Double
        '    Dim [Step] As Double
        '    Dim DBObjectCollection As New DBObjectCollection
        '    Dim StartPoint As Point3d
        '    Dim EndPoint As Point3d
        '    Dim Line As Curve

        '    Point3dcollection = New Point3dCollection()
        '    Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
        '    [Step] = (Lenght / 250)
        '    For CurvePosition As Double = 0 To Lenght Step [Step]
        '        Point3d = Curve.GetPointAtDist(CurvePosition)
        '        With PromptNestedEntityOptions
        '            .NonInteractivePickPoint = Point3d
        '            .UseNonInteractivePickPoint = True
        '        End With
        '        Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
        '    Next

        '    For Index As Integer = 0 To Point3dcollection.Count - 1 Step 2
        '        StartPoint = Point3dcollection(Index)
        '        EndPoint = Point3dcollection(Index + 1)
        '        If IsNothing(TypedValueCollection) = False Then
        '            PromptSelectionResult = Editor.SelectFence(New Point3dCollection({StartPoint, EndPoint}), New SelectionFilter(TypedValueCollection.ToArray))
        '        Else
        '            PromptSelectionResult = Editor.SelectFence(New Point3dCollection({StartPoint, EndPoint}))
        '        End If
        '        If IsNothing(PromptSelectionResult) = False AndAlso PromptSelectionResult.Status = PromptStatus.OK Then
        '            For Each x As FenceSelectedObject In PromptSelectionResult.Value
        '                If x.ObjectId.Equals(Curve.ObjectId) = False Then
        '                    If Elements.ContainsKey(x.ObjectId) = False Then
        '                        Elements.Add(x.ObjectId, New List(Of Point3dCollection)({New Point3dCollection({StartPoint, EndPoint})}))
        '                    Else
        '                        Elements(x.ObjectId).AddRange({New Point3dCollection({StartPoint, EndPoint})})
        '                    End If
        '                End If
        '            Next
        '        End If
        '    Next

        '    For Each Element As KeyValuePair(Of ObjectId, List(Of Point3dCollection)) In Elements
        '        For Each Points As Point3dCollection In Element.Value
        '            For Index As Integer = 0 To Points.Count - 1 Step 2
        '                StartPoint = Points(Index)
        '                EndPoint = Points(Index + 1)
        '                Line = New Line(StartPoint, EndPoint)
        '                Lenght = Line.GetDistanceAtParameter(Line.EndParam) - Line.GetDistanceAtParameter(Line.StartParam)
        '                For CurvePosition As Double = 0 To Lenght Step Tolerance.Global.EqualPoint + 0.000000001
        '                    Point3d = Line.GetPointAtDist(CurvePosition)
        '                    With PromptNestedEntityOptions
        '                        .NonInteractivePickPoint = Point3d
        '                        .UseNonInteractivePickPoint = True
        '                    End With
        '                    PromptNestedEntityResult = Editor.GetNestedEntity(PromptNestedEntityOptions)
        '                    If IsNothing(PromptNestedEntityResult) = False AndAlso PromptNestedEntityResult.Status = PromptStatus.OK Then
        '                        If PromptNestedEntityResult.ObjectId.Equals(Curve.Id) = False Then
        '                            If Result.ContainsKey(PromptNestedEntityResult.ObjectId) = False Then
        '                                Result.Add(PromptNestedEntityResult.ObjectId, New Point3dCollection({PromptNestedEntityResult.PickedPoint}))
        '                            Else
        '                                Result(PromptNestedEntityResult.ObjectId).Add(PromptNestedEntityResult.PickedPoint)
        '                            End If
        '                        End If
        '                    End If
        '                Next
        '            Next
        '        Next
        '    Next

        '    'Retorno
        '    Return Result

        'End Function

        ''' <summary>
        ''' Seleciona entidades com base em um conjunto de pontos (polígono virtual) 
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos</param>
        ''' <param name="SelectReachMode">Modo de seleção</param>
        ''' <param name="TypedValueCollection"></param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectFramePoints(Point3dCollection As Point3dCollection, SelectReachMode As eSelectReachMode, Optional TypedValueCollection As List(Of TypedValue) = Nothing) As Object
            Try
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim PromptSelectionResult As PromptSelectionResult = Nothing
                Dim FenceResult As ObjectIdCollection = Nothing
                Dim StartPoint As Point3d
                Dim EndPoint As Point3d
                Select Case SelectReachMode
                    Case eSelectReachMode.SelectCrossingPolygon
                        If IsNothing(TypedValueCollection) = False Then
                            PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dCollection, New SelectionFilter(TypedValueCollection.ToArray))
                        Else
                            PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dCollection)
                        End If
                    Case eSelectReachMode.SelectWindowPolygon
                        If IsNothing(TypedValueCollection) = False Then
                            PromptSelectionResult = Editor.SelectWindowPolygon(Point3dCollection, New SelectionFilter(TypedValueCollection.ToArray))
                        Else
                            PromptSelectionResult = Editor.SelectWindowPolygon(Point3dCollection)
                        End If
                    Case eSelectReachMode.SelectFence
                        If IsNothing(TypedValueCollection) = False Then
                            FenceResult = New ObjectIdCollection
                            For Index As Integer = 0 To (Point3dCollection.Count - 1)
                                StartPoint = Point3dCollection.Item(Index)
                                If Index <> (Point3dCollection.Count - 1) Then
                                    EndPoint = Point3dCollection.Item(Index + 1)
                                Else
                                    EndPoint = Point3dCollection.Item(0)
                                End If
                                PromptSelectionResult = Editor.SelectFence(New Point3dCollection({StartPoint, EndPoint}), New SelectionFilter(TypedValueCollection.ToArray))
                                Select Case PromptSelectionResult.Status
                                    Case PromptStatus.OK
                                        For Each ObjectId As ObjectId In PromptSelectionResult.Value.GetObjectIds
                                            FenceResult.Add(ObjectId)
                                        Next
                                End Select
                            Next
                        Else
                            PromptSelectionResult = Editor.SelectFence(Point3dCollection)
                        End If
                End Select
                Select Case SelectReachMode
                    Case eSelectReachMode.SelectFence
                        If FenceResult.Count = 0 Then
                            FenceResult = New ObjectIdCollection
                        End If
                        Return FenceResult
                    Case Else
                        Select Case PromptSelectionResult.Status
                            Case PromptStatus.OK
                                Return New ObjectIdCollection(PromptSelectionResult.Value.GetObjectIds)
                            Case Else
                                Return New ObjectIdCollection
                        End Select
                End Select

            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Seleciona entidades com base em um conjunto de pontos (polígono virtual) 
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos</param>
        ''' <param name="SelectReachMode">Modo de seleção</param>
        ''' <param name="TypedValueCollection"></param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectFramePoints(Point3dCollection() As Point3d, SelectReachMode As eSelectReachMode, Optional TypedValueCollection As List(Of TypedValue) = Nothing) As Object
            Return Engine2.AcadInterface.SelectFramePoints(New Point3dCollection(Point3dCollection), SelectReachMode, TypedValueCollection)
        End Function

        ''' <summary>
        ''' Seleciona entidades com base em um conjunto de pontos (polígono virtual) 
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos</param>
        ''' <param name="SelectReachMode">Modo de seleção</param>
        ''' <param name="TypedValueCollection"></param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectFramePoints(Point3dCollection As List(Of Point3d), SelectReachMode As eSelectReachMode, Optional TypedValueCollection As List(Of TypedValue) = Nothing) As Object
            Return Engine2.AcadInterface.SelectFramePoints(New Point3dCollection(Point3dCollection.ToArray), SelectReachMode, TypedValueCollection)
        End Function

        ''' <summary>
        ''' Seleção de valor angular 
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="BasePoint">Ponto de base</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita qualquer valor</param>
        ''' <param name="AllowNone">Determina se aceita enter direto</param>
        ''' <param name="AllowZero">Determina se aceita zero</param>
        ''' <param name="UseDashedLine">Determina se deve usar linha tracejada</param>
        ''' <param name="UseAngleBase">Determina se será usado o ângulo de base</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional BasePoint As Object = Nothing, Optional AllowArbitraryInput As Boolean = False, Optional AllowNone As Boolean = False, Optional AllowZero As Boolean = True, Optional UseDashedLine As Boolean = False, Optional UseAngleBase As Boolean = False, Optional ExceptionMessages As Boolean = False) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptAngleOptions As New PromptAngleOptions(vbLf & Message)
            With PromptAngleOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                    .UseDefaultValue = True
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                If IsNothing(BasePoint) = False Then
                    .BasePoint = BasePoint
                    .UseBasePoint = True
                Else
                    .BasePoint = Nothing
                    .UseBasePoint = False
                End If
                .AllowArbitraryInput = AllowArbitraryInput
                .AllowNone = AllowNone
                .AllowZero = AllowZero
                .UseDashedLine = UseDashedLine
                .UseAngleBase = UseAngleBase
            End With
            Dim PromptDoubleResult As PromptDoubleResult
            Document.Window.Focus()
            PromptDoubleResult = Editor.GetAngle(PromptAngleOptions)
            Select Case PromptDoubleResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptDoubleResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return PromptDoubleResult.Value
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de entidade
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="RejectMessage">Mensagem de rejeição</param>
        ''' <param name="AllowObjectOnLockedLayer">Determina se a seleção de entidades em camadas bloqueadas serão aceitas</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <param name="AllowedClassCollection">Determina os tipos de objetos que serão aceitos</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing/Entity/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetEntity(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional RejectMessage As String = "Seleção inválida.", Optional AllowObjectOnLockedLayer As Boolean = False, Optional AllowNone As Boolean = False, Optional AllowedClassCollection As List(Of System.Type) = Nothing, Optional ExceptionMessages As Boolean = False) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptEntityOptions As New PromptEntityOptions(vbLf & Message)
            With PromptEntityOptions
                .AllowObjectOnLockedLayer = AllowObjectOnLockedLayer
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                .AllowNone = AllowNone
                .SetRejectMessage(vbLf & RejectMessage & vbLf)
                If IsNothing(AllowedClassCollection) = False Then
                    For Each [Type] As System.Type In AllowedClassCollection
                        .AddAllowedClass([Type], False)
                    Next
                End If
            End With
            Dim PromptEntityResult As PromptEntityResult = Nothing
            Document.Window.Focus()
            PromptEntityResult = Editor.GetEntity(PromptEntityOptions)
            Select Case PromptEntityResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptEntityResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return Engine2.ConvertObject.ObjectIDToEntity(PromptEntityResult.ObjectId)
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de entidade
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="RejectMessage">Mensagem de rejeição</param>
        ''' <param name="AllowObjectOnLockedLayer">Determina se a seleção de entidades em camadas bloqueadas serão aceitas</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <param name="AllowedClassCollection">Determina os tipos de objetos que serão aceitos</param>
        ''' <returns>PromptEntityResult</returns>
        ''' <remarks></remarks>
        Public Shared Function GetEntity2(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional RejectMessage As String = "Seleção inválida.", Optional AllowObjectOnLockedLayer As Boolean = False, Optional AllowNone As Boolean = False, Optional AllowedClassCollection As List(Of System.Type) = Nothing) As PromptEntityResult
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptEntityOptions As New PromptEntityOptions(vbLf & Message)
            With PromptEntityOptions
                .AllowObjectOnLockedLayer = AllowObjectOnLockedLayer
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                .AllowNone = AllowNone
                .SetRejectMessage(vbLf & RejectMessage & vbLf)
                If IsNothing(AllowedClassCollection) = False Then
                    For Each [Type] As System.Type In AllowedClassCollection
                        .AddAllowedClass([Type], False)
                    Next
                End If
            End With
            Document.Window.Focus()
            Return Editor.GetEntity(PromptEntityOptions)
        End Function

        ''' <summary>
        '''  Seleção de múltiplas entidades
        ''' </summary>
        ''' <param name="MessageForAdding">Mensagem para inclusão de seleção</param>
        ''' <param name="MessageForRemoval">Mensagem para remoção de seleção</param>
        ''' <param name="TypedValueCollection">Filtro de seleção</param>
        ''' <param name="RejectObjectsOnLockedLayers">Determina se a seleção de entidades em camadas bloqueadas serão aceitas</param>
        ''' <param name="AllowDuplicates">Determina se duplicações serão aceitas</param>
        ''' <param name="AllowSubSelections">Determina se sub-seleção será aceita</param>
        ''' <param name="ForceSubSelections">Determina se a seleção deve conter sub seleção</param>
        ''' <param name="RejectPaperspaceViewport">Determina se a seleção irá considerar entidades fora do espaço do modelo</param>
        ''' <param name="RejectObjectsFromNonCurrentSpace">Determina se a seleção irá considerar seleção fora do espaço corrente</param>
        ''' <param name="SelectEverythingInAperture">Determina se a seleção permite o conteúdo de aperture</param>
        ''' <param name="SingleOnly">Determina se será habilitada a seleção única</param>
        ''' <param name="SinglePickInSpace">Determina se será habilitada o ponto único na seleção</param>
        ''' <param name="PrepareOptionalDetails">Determina se serão preparados os detalhes opcionais</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing\ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function GetSelection(Optional MessageForAdding As Object = Nothing, Optional MessageForRemoval As Object = Nothing, Optional TypedValueCollection As List(Of TypedValue) = Nothing, Optional RejectObjectsOnLockedLayers As Boolean = True, Optional AllowDuplicates As Boolean = False, Optional AllowSubSelections As Boolean = False, Optional ForceSubSelections As Boolean = False, Optional RejectPaperspaceViewport As Boolean = False, Optional RejectObjectsFromNonCurrentSpace As Boolean = False, Optional SelectEverythingInAperture As Boolean = False, Optional SingleOnly As Boolean = False, Optional SinglePickInSpace As Boolean = False, Optional PrepareOptionalDetails As Boolean = False, Optional ExceptionMessages As Boolean = False) As ObjectIdCollection
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptSelectionOptions As PromptSelectionOptions = New PromptSelectionOptions()
            With PromptSelectionOptions
                .RejectObjectsOnLockedLayers = RejectObjectsOnLockedLayers
                .AllowDuplicates = AllowDuplicates
                .AllowSubSelections = AllowSubSelections
                .ForceSubSelections = ForceSubSelections
                .RejectPaperspaceViewport = RejectPaperspaceViewport
                .RejectObjectsFromNonCurrentSpace = RejectObjectsFromNonCurrentSpace
                .SelectEverythingInAperture = SelectEverythingInAperture
                .SingleOnly = SingleOnly
                .SinglePickInSpace = SinglePickInSpace
                .PrepareOptionalDetails = PrepareOptionalDetails
                If IsNothing(MessageForAdding) = False Then
                    .MessageForAdding = Convert.ToString(vbLf & MessageForAdding)
                End If
                If IsNothing(MessageForRemoval) = False Then
                    .MessageForRemoval = Convert.ToString(vbLf & MessageForRemoval)
                End If
            End With
            Dim PromptSelectionResult As PromptSelectionResult = Nothing
            Document.Window.Focus()
            If IsNothing(TypedValueCollection) = False Then
                PromptSelectionResult = Editor.GetSelection(PromptSelectionOptions, New SelectionFilter(TypedValueCollection.ToArray))
            Else
                PromptSelectionResult = Editor.GetSelection(PromptSelectionOptions)
            End If
            Select Case PromptSelectionResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return Nothing
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return New ObjectIdCollection(PromptSelectionResult.Value.GetObjectIds)
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de sub entidade
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="NonInteractivePickPoint">Ponto que não interage com a seleção</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing/ObjectID/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetNestedEntity(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional NonInteractivePickPoint As Object = Nothing, Optional AllowNone As Boolean = False, Optional ExceptionMessages As Boolean = False) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptNestedEntityOptions As New PromptNestedEntityOptions(vbLf & Message)
            With PromptNestedEntityOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                If IsNothing(NonInteractivePickPoint) = False Then
                    .NonInteractivePickPoint = NonInteractivePickPoint
                    .UseNonInteractivePickPoint = True
                Else
                    .UseNonInteractivePickPoint = False
                End If
                .AllowNone = AllowNone
            End With
            Dim PromptNestedEntityResult As PromptNestedEntityResult
            Document.Window.Focus()
            PromptNestedEntityResult = Editor.GetNestedEntity(PromptNestedEntityOptions)
            Select Case PromptNestedEntityResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptNestedEntityResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return Engine2.ConvertObject.ObjectIDToEntity(PromptNestedEntityResult.ObjectId)
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de sub entidade
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="NonInteractivePickPoint">Ponto que não interage com a seleção</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <returns>PromptNestedEntityResult</returns>
        ''' <remarks></remarks>
        Public Shared Function GetNestedEntity2(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional NonInteractivePickPoint As Object = Nothing, Optional AllowNone As Boolean = False) As PromptNestedEntityResult
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptNestedEntityOptions As New PromptNestedEntityOptions(vbLf & Message)
            With PromptNestedEntityOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                If IsNothing(NonInteractivePickPoint) = False Then
                    .NonInteractivePickPoint = NonInteractivePickPoint
                    .UseNonInteractivePickPoint = True
                Else
                    .UseNonInteractivePickPoint = False
                End If
                .AllowNone = AllowNone
            End With
            Document.Window.Focus()
            Return Editor.GetNestedEntity(PromptNestedEntityOptions)
        End Function

        ''' <summary>
        ''' Seleciona todas as entidades
        ''' </summary>
        ''' <param name="TypedValueCollection">Filtro de seleção</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectAll(Optional TypedValueCollection As List(Of TypedValue) = Nothing) As ObjectIdCollection
            SelectAll = New ObjectIdCollection
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim PromptSelectionResult As PromptSelectionResult
            If IsNothing(TypedValueCollection) = False Then
                PromptSelectionResult = Editor.SelectAll(New SelectionFilter(TypedValueCollection.ToArray))
            Else
                PromptSelectionResult = Editor.SelectAll
            End If
            If PromptSelectionResult.Status = PromptStatus.OK Then
                SelectAll = New ObjectIdCollection(PromptSelectionResult.Value.GetObjectIds)
            End If
            Return SelectAll
        End Function

        ''' <summary>
        ''' Solicita a entrada de um texto
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="DefaltValue">Opção padrão</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <param name="AllowSpaces">Determina se espaços serão aceitos</param>
        ''' <returns>Nothing/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetString(Message As String, Optional DefaltValue As String = "", Optional ExceptionMessages As Boolean = False, Optional AllowSpaces As Boolean = False) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptStringOptions As New PromptStringOptions(vbLf & Message)
            With PromptStringOptions
                If DefaltValue.Trim <> "" Then
                    .UseDefaultValue = True
                    .AppendKeywordsToMessage = True
                    .DefaultValue = DefaltValue
                End If
                .AllowSpaces = AllowSpaces
            End With
            Dim PromptResult As PromptResult = Nothing
            Document.Window.Focus()
            PromptResult = Editor.GetString(PromptStringOptions)
            Select Case PromptResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return PromptResult.StringResult
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Solicita a entrada de seleção do usuário
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Coleção de opções</param>
        ''' <param name="KeywordsDefalt">Opção padrão</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita valores diversos</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param> 
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetKeywords(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional AllowArbitraryInput As Boolean = False, Optional AllowNone As Boolean = False, Optional ExceptionMessages As Boolean = False) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptKeywordOptions As New PromptKeywordOptions(vbLf & Message)
            With PromptKeywordOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                .AllowNone = AllowNone
                .AllowArbitraryInput = AllowArbitraryInput
            End With
            Dim PromptResult As PromptResult = Nothing
            Document.Window.Focus()
            PromptResult = Editor.GetKeywords(PromptKeywordOptions)
            Select Case PromptResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return PromptResult.StringResult
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de ponto
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="BasePoint">Ponto de base</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita valores diversos</param>
        ''' <param name="UseDashedLine">Determina se deve usar linha tracejada</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing/Point3d/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetPoint(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional BasePoint As Object = Nothing, Optional AllowArbitraryInput As Boolean = False, Optional UseDashedLine As Boolean = False, Optional AllowNone As Boolean = False, Optional ExceptionMessages As Boolean = False) As Object
            If IsNothing(BasePoint) = False Then
                If BasePoint.GetType.Name <> "Point3d" Then
                    Throw New System.Exception("O formato do parâmetro BasePoint não é válido, somente Point3d é aceito.")
                End If
            End If
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptPointOptions As New PromptPointOptions(vbLf & Message)
            Dim PromptPointResult As PromptPointResult
            With PromptPointOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                If IsNothing(BasePoint) = False Then
                    .BasePoint = BasePoint
                    .UseBasePoint = True
                Else
                    .BasePoint = Nothing
                    .UseBasePoint = False
                End If
                .AllowArbitraryInput = AllowArbitraryInput
                .UseDashedLine = UseDashedLine
                .AllowNone = AllowNone
            End With
            Document.Window.Focus()
            PromptPointResult = Editor.GetPoint(PromptPointOptions)
            Select Case PromptPointResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptPointResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return PromptPointResult.Value
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Seleção de ponto de canto
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Atalho padrão</param>
        ''' <param name="BasePoint">Ponto de base</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita valores diversos</param>
        ''' <param name="UseDashedLine">Determina se deve usar linha tracejada</param>
        ''' <param name="AllowNone">Determina se nenhuma seleção é aceita</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <returns>Nothing/Point3d/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCorner(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional BasePoint As Object = Nothing, Optional AllowArbitraryInput As Boolean = False, Optional UseDashedLine As Boolean = False, Optional AllowNone As Boolean = False, Optional ExceptionMessages As Boolean = False) As Object
            If IsNothing(BasePoint) = False Then
                If BasePoint.GetType.Name <> "Point3d" Then
                    Throw New System.Exception("O formato do parâmetro BasePoint não é válido, somente Point3d é aceito.")
                End If
            End If
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptCornerOptions As New PromptCornerOptions(vbLf & Message, BasePoint)
            Dim PromptPointResult As PromptPointResult
            With PromptCornerOptions
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    .Keywords.Default = KeywordsDefalt
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
                .AllowArbitraryInput = AllowArbitraryInput
                .UseDashedLine = UseDashedLine
                .AllowNone = AllowNone
            End With
            Document.Window.Focus()
            PromptPointResult = Editor.GetCorner(PromptCornerOptions)
            Select Case PromptPointResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptPointResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    Return PromptPointResult.Value
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Solicita um número real
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Valor padrão</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita qualquer valor</param>
        ''' <param name="AllowNegative">Determina se aceita negativo</param>
        ''' <param name="AllowNone">Determina se aceita enter direto</param>
        ''' <param name="AllowZero">Determina se aceita zero</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <param name="MinValue">Valor mínimo</param>
        ''' <param name="MaxValue">Valor máximo</param>
        ''' <returns>Nothing/Double/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetDouble(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional AllowArbitraryInput As Boolean = False, Optional AllowNegative As Boolean = False, Optional AllowNone As Boolean = False, Optional AllowZero As Boolean = False, Optional ExceptionMessages As Boolean = False, Optional MinValue As Object = Nothing, Optional MaxValue As Object = Nothing) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptDoubleOptions As New PromptDoubleOptions(vbLf & Message)
            With PromptDoubleOptions
                .AllowArbitraryInput = AllowArbitraryInput
                .AllowNegative = AllowNegative
                .AllowNone = AllowNone
                .AllowZero = AllowZero
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    If Keywords.Contains(KeywordsDefalt) = True Then
                        .Keywords.Default = KeywordsDefalt
                    Else
                        .DefaultValue = KeywordsDefalt
                    End If
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
            End With
            Dim PromptDoubleResult As PromptDoubleResult
            Document.Window.Focus()
            PromptDoubleResult = Editor.GetDouble(PromptDoubleOptions)
            Select Case PromptDoubleResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptDoubleResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    If IsNothing(MinValue) = False And IsNothing(MaxValue) = False Then
                        If CDbl(PromptDoubleResult.Value) >= CDbl(MinValue) And CDbl(PromptDoubleResult.Value) <= CDbl(MaxValue) Then
                            Return PromptDoubleResult.Value
                        Else
                            If ExceptionMessages = True Then
                                Editor.WriteMessage(vbLf & "Valor inválido, o valor informado deve estar entre '" & MinValue.ToString & "' e '" & MaxValue.ToString & "'." & vbLf)
                            End If
                            Return Nothing
                        End If
                    Else
                        Return PromptDoubleResult.Value
                    End If
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Solicita um número inteiro
        ''' </summary>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Keywords">Atalhos</param>
        ''' <param name="KeywordsDefalt">Valor padrão</param>
        ''' <param name="AllowArbitraryInput">Determina se aceita qualquer valor</param>
        ''' <param name="AllowNegative">Determina se aceita negativo</param>
        ''' <param name="AllowNone">Determina se aceita enter direto</param>
        ''' <param name="AllowZero">Determina se aceita zero</param>
        ''' <param name="ExceptionMessages">Determina se o comando irá gerar mensagens de cancelamento ou erro</param>
        ''' <param name="MinValue">Valor mínimo</param>
        ''' <param name="MaxValue">Valor máximo</param>
        ''' <returns>Nothing/Integer/String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetInteger(Message As String, Optional Keywords As ArrayList = Nothing, Optional KeywordsDefalt As String = "", Optional AllowArbitraryInput As Boolean = False, Optional AllowNegative As Boolean = False, Optional AllowNone As Boolean = False, Optional AllowZero As Boolean = False, Optional ExceptionMessages As Boolean = False, Optional MinValue As Object = Nothing, Optional MaxValue As Object = Nothing) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PromptIntegerOptions As New PromptIntegerOptions(vbLf & Message)
            With PromptIntegerOptions
                .AllowArbitraryInput = AllowArbitraryInput
                .AllowNegative = AllowNegative
                .AllowNone = AllowNone
                .AllowZero = AllowZero
                If IsNothing(Keywords) = False Then
                    For Each Keyword As String In Keywords
                        .Keywords.Add(Keyword)
                    Next
                End If
                If KeywordsDefalt.Trim <> "" Then
                    If Keywords.Contains(KeywordsDefalt) = True Then
                        .Keywords.Default = KeywordsDefalt
                    Else
                        .DefaultValue = KeywordsDefalt
                    End If
                End If
                If .Keywords.Count > 0 Then
                    .AppendKeywordsToMessage = True
                Else
                    .AppendKeywordsToMessage = False
                End If
            End With
            Dim PromptIntegerResult As PromptIntegerResult
            Document.Window.Focus()
            PromptIntegerResult = Editor.GetInteger(PromptIntegerOptions)
            Select Case PromptIntegerResult.Status
                Case PromptStatus.Cancel
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Comando cancelado." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Error
                    If ExceptionMessages = True Then
                        Editor.WriteMessage(vbLf & "Erro no processo de seleção." & vbLf)
                    End If
                    Return Nothing
                Case PromptStatus.Keyword
                    Return PromptIntegerResult.StringResult
                Case PromptStatus.Modeless
                    Return Nothing
                Case PromptStatus.None
                    Return Nothing
                Case PromptStatus.OK
                    If IsNothing(MinValue) = False And IsNothing(MaxValue) = False Then
                        If CInt(PromptIntegerResult.Value) >= CInt(MinValue) And CInt(PromptIntegerResult.Value) <= CInt(MaxValue) Then
                            Return PromptIntegerResult.Value
                        Else
                            If ExceptionMessages = True Then
                                Editor.WriteMessage(vbLf & "Valor inválido, o valor informado deve estar entre '" & MinValue.ToString & "' e '" & MaxValue.ToString & "'." & vbLf)
                            End If
                            Return Nothing
                        End If
                    Else
                        Return PromptIntegerResult.Value
                    End If
                Case PromptStatus.Other
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Retorna a coleção de entidades contidas em um limite de alcance definido pela geometria de uma entidade
        ''' </summary>
        ''' <param name="Curve">Entidade</param>
        ''' <param name="TypedValueCollection">Filtro de seleção</param>
        ''' <param name="SelectReachMode">Modo de seleção</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <param name="ExcludeCurve">Determina se a entidade de origem (Curve) será excluída do resultado (Apenas se a Curva for real)</param>
        ''' <param name="IgnoreIds">Os Id´s informados serão ignorados pelo retorno</param>
        ''' <param name="StepSurfaceCurve">Número de segmentos ao longo da entidade para detecção de entidades na curva</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectReach(Curve As Curve, Optional TypedValueCollection As List(Of TypedValue) = Nothing, Optional SelectReachMode As eSelectReachMode = eSelectReachMode.SelectFence, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False, Optional ExcludeCurve As Boolean = False, Optional IgnoreIds As ObjectIdCollection = Nothing, Optional StepSurfaceCurve As Integer = 250) As DBObjectCollection
            SelectReach = New DBObjectCollection
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Point3dcollection As Point3dCollection
            Dim Point3d As Point3d
            Dim Lenght As Double
            Dim [Step] As Double
            Dim PromptSelectionResult As PromptSelectionResult = Nothing
            Dim DBObjectCollection As New DBObjectCollection
            Dim ObjectIdCollection As ObjectIdCollection
            Dim Entity As Entity
            'Dim TypedValueClassNames As ArrayList = Engine2.AcadInterface.GetClassNames(TypedValueCollection)
            Point3dcollection = New Point3dCollection()
            If Curve.Closed = True Then
                Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
                [Step] = (Lenght / StepSurfaceCurve)
                For CurvePosition As Double = 0 To Lenght Step [Step]
                    Point3d = Curve.GetPointAtDist(CurvePosition)
                    Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
                Next
            Else
                Point3dcollection.Add(Curve.StartPoint)
                Point3dcollection.Add(Curve.EndPoint)
            End If
            Select Case SelectReachMode
                Case eSelectReachMode.SelectCrossingPolygon
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dcollection)
                    End If
                Case eSelectReachMode.SelectWindowPolygon
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectWindowPolygon(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectWindowPolygon(Point3dcollection)
                    End If
                Case eSelectReachMode.SelectFence
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectFence(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectFence(Point3dcollection)
                    End If
            End Select
            Select Case PromptSelectionResult.Status
                Case PromptStatus.OK
                    If SubItems = False Then
                        For Each ObjectId As ObjectId In PromptSelectionResult.Value.GetObjectIds
                            Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
                            If IsNothing(Entity) = False Then
                                If IsNothing(IgnoreIds) = False Then
                                    If IgnoreIds.Contains(Entity.ObjectId) = False Then
                                        'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                        SelectReach.Add(Entity)
                                        'End If
                                    End If
                                Else
                                    'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                    SelectReach.Add(Entity)
                                    'End If
                                End If
                            End If
                        Next
                    Else
                        DBObjectCollection = New DBObjectCollection
                        ObjectIdCollection = New ObjectIdCollection(PromptSelectionResult.Value.GetObjectIds)
                        While ObjectIdCollection.Count <> 0
                            Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectIdCollection.Item(0))
                            If IsNothing(Entity) = False Then
                                If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                    DBObjectCollection = Engine2.EntityInteration.GetSubEntitys(Entity)
                                    For Each DBObject As DBObject In DBObjectCollection
                                        Entity = Engine2.ConvertObject.DBObjectToEntity(DBObject)
                                        If NestedEntitys = True Then
                                            If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                                ObjectIdCollection.Add(Entity.ObjectId)
                                            Else
                                                If IsNothing(IgnoreIds) = False Then
                                                    If IgnoreIds.Contains(Entity.ObjectId) = False Then
                                                        'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                                        SelectReach.Add(Entity)
                                                        'End If
                                                    End If
                                                Else
                                                    'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                                    SelectReach.Add(Entity)
                                                    'End If
                                                End If
                                            End If
                                        Else
                                            If IsNothing(IgnoreIds) = False Then
                                                If IgnoreIds.Contains(Entity.ObjectId) = False Then
                                                    'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                                    SelectReach.Add(Entity)
                                                    'End If
                                                End If
                                            Else
                                                'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                                SelectReach.Add(Entity)
                                                'End If
                                            End If
                                        End If
                                    Next
                                Else
                                    If IsNothing(IgnoreIds) = False Then
                                        If IgnoreIds.Contains(Entity.ObjectId) = False Then
                                            'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                            SelectReach.Add(Entity)
                                            'End If
                                        End If
                                    Else
                                        'If TypedValueClassNames.Contains(Entity.GetType.Name) = True Then
                                        SelectReach.Add(Entity)
                                        'End If
                                    End If
                                End If
                            End If
                            ObjectIdCollection.RemoveAt(0)
                        End While
                    End If
                    If SelectReach.Contains(Curve) = True And ExcludeCurve = True Then
                        SelectReach.Remove(Curve)
                    End If
                    Return SelectReach
                Case Else
                    Return SelectReach
            End Select
        End Function

        ''' <summary>
        ''' Retorna a coleção de entidades contidas em um limite de alcance a partir de um ponto
        ''' </summary>
        ''' <param name="Center">Centro da seleção</param>
        ''' <param name="Reach">Alcance da seleção</param>
        ''' <param name="TypedValueCollection">Filtro de seleção</param>
        ''' <param name="SelectReachMode">Modo de seleção</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <param name="IgnoreIds">Os Id´s informados serão ignorados pelo retorno</param>
        ''' <param name="StepSurfaceCurve">Número de segmentos ao longo da entidade para detecção de entidades na curva</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectReach(Center As Point3d, Reach As Double, Optional TypedValueCollection As List(Of TypedValue) = Nothing, Optional SelectReachMode As eSelectReachMode = eSelectReachMode.SelectFence, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False, Optional IgnoreIds As ObjectIdCollection = Nothing, Optional StepSurfaceCurve As Integer = 250) As DBObjectCollection
            SelectReach = New DBObjectCollection
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Curve As Curve
            Dim Point3dcollection As Point3dCollection
            Dim Point3d As Point3d
            Dim Lenght As Double
            Dim [Step] As Double
            Dim PromptSelectionResult As PromptSelectionResult = Nothing
            Dim Circle As Circle
            Dim DBObjectCollection As New DBObjectCollection
            Dim Entity As Entity
            Dim ObjectIdCollection As ObjectIdCollection
            'Dim TypedValueClassNames As ArrayList = Engine2.AcadInterface.GetClassNames(TypedValueCollection)
            Circle = New Circle(Center, Vector3d.ZAxis, Reach)
            Curve = Circle
            Point3dcollection = New Point3dCollection()
            Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
            [Step] = (Lenght / StepSurfaceCurve)
            For CurvePosition As Double = 0 To Lenght Step [Step]
                Point3d = Curve.GetPointAtDist(CurvePosition)
                Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
            Next
            Select Case SelectReachMode
                Case eSelectReachMode.SelectCrossingPolygon
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectCrossingPolygon(Point3dcollection)
                    End If
                Case eSelectReachMode.SelectWindowPolygon
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectWindowPolygon(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectWindowPolygon(Point3dcollection)
                    End If
                Case eSelectReachMode.SelectFence
                    If IsNothing(TypedValueCollection) = False Then
                        PromptSelectionResult = Editor.SelectFence(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                    Else
                        PromptSelectionResult = Editor.SelectFence(Point3dcollection)
                    End If
            End Select
            Select Case PromptSelectionResult.Status
                Case PromptStatus.OK
                    If SubItems = False Then
                        For Each ObjectId As ObjectId In PromptSelectionResult.Value.GetObjectIds
                            Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
                            If IsNothing(Entity) = False Then
                                If IsNothing(IgnoreIds) = False Then
                                    If IgnoreIds.Contains(Entity.Id) = False Then
                                        SelectReach.Add(Entity)
                                    End If
                                Else
                                    SelectReach.Add(Entity)
                                End If
                            End If
                        Next
                    Else
                        DBObjectCollection = New DBObjectCollection
                        ObjectIdCollection = New ObjectIdCollection(PromptSelectionResult.Value.GetObjectIds)
                        While ObjectIdCollection.Count <> 0
                            Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectIdCollection.Item(0))
                            If IsNothing(Entity) = False Then
                                If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                    DBObjectCollection = Engine2.EntityInteration.GetSubEntitys(Entity)
                                    For Each DBObject As DBObject In DBObjectCollection
                                        Entity = Engine2.ConvertObject.DBObjectToEntity(DBObject)
                                        If NestedEntitys = True Then
                                            If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                                ObjectIdCollection.Add(Entity.ObjectId)
                                            Else
                                                If IsNothing(IgnoreIds) = False Then
                                                    If IgnoreIds.Contains(Entity.Id) = False Then
                                                        SelectReach.Add(Entity)
                                                    End If
                                                Else
                                                    SelectReach.Add(Entity)
                                                End If
                                            End If
                                        Else
                                            If IsNothing(IgnoreIds) = False Then
                                                If IgnoreIds.Contains(Entity.Id) = False Then
                                                    SelectReach.Add(Entity)
                                                End If
                                            Else
                                                SelectReach.Add(Entity)
                                            End If
                                        End If
                                    Next
                                Else
                                    If IsNothing(IgnoreIds) = False Then
                                        If IgnoreIds.Contains(Entity.Id) = False Then
                                            SelectReach.Add(Entity)
                                        End If
                                    Else
                                        SelectReach.Add(Entity)
                                    End If
                                End If
                            End If
                            ObjectIdCollection.RemoveAt(0)
                        End While
                    End If
                    Return SelectReach
                Case Else
                    Return SelectReach
            End Select
        End Function

#Region "SelectAtPoint" ' Repõe a função SelectAtPoint existente no VBA ´mas não existente no .Net

        ''' <summary>
        ''' Operadores para filtragem de XData
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eXDataOperator
            None = 0
            _AND = 1
            _OR = 2
        End Enum

        ''' <summary>
        ''' Seleciona as entidades contidas em um ponto (Método de pesquisa com base no banco de dados do desenho. Seleciona em toda a área gráfica mesmo que não visível)
        ''' </summary>
        ''' <param name="Point3d">Ponto a ser pesquisado</param>
        ''' <param name="AllowedClassCollection">Classes com os tipos de entidades a serem observadas</param>
        ''' <param name="XDataKeys">Chaves XData a serem procuradas</param>
        ''' <param name="IgnoreIDs">ID´s a serem ignorados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SelectAtPoint(Point3d As Point3d, AllowedClassCollection As List(Of System.Type), Optional XDataOperator As eXDataOperator = eXDataOperator.None, Optional XDataKeys As List(Of String) = Nothing, Optional IgnoreIDs As ObjectIdCollection = Nothing) As ObjectIdCollection
            SelectAtPoint = New ObjectIdCollection
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim DBObject As DBObject
            Dim DBObjectCollection As New DBObjectCollection
            Dim AllowedClassCollection2 As New List(Of System.Type)
            Dim Valid As Boolean = False
            Dim Count As Integer = 0
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForRead)
                    For Each ObjectId As ObjectId In BlockTableRecord
                        DBObject = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                        If IsNothing(XDataKeys) = False And IsNothing(DBObject.XData) = False Then
                            Select Case XDataOperator
                                Case eXDataOperator._AND
                                    For Each XDataKey As String In XDataKeys
                                        If DBObject.ContainsXData(XDataKey) = True Then
                                            Count = (Count + 1)
                                        End If
                                    Next
                                    If Count.Equals(XDataKeys.Count) Then
                                        Valid = True
                                    Else
                                        Valid = False
                                    End If
                                Case eXDataOperator._OR
                                    For Each XDataKey As String In XDataKeys
                                        If DBObject.ContainsXData(XDataKey) = True Then
                                            Valid = True
                                            Exit For
                                        End If
                                    Next
                                Case Else
                                    Throw New System.Exception("Parâmetro XDataOperator não informado.")
                            End Select
                        Else
                            Valid = True
                        End If
                         If AllowedClassCollection.Contains(DBObject.GetType) = True And Valid = True And If(IsNothing(IgnoreIDs) = False, If(IgnoreIDs.Contains(DBObject.ObjectId) = True, 0 = 1, 1 = 1), 1 = 1) Then
                            If Engine2.EntityInteration.IsInsidePoint(DBObject, Point3d) = True Then
                                SelectAtPoint.Add(ObjectId)
                            End If
                        End If
                    Next
                End Using
            End Using
            Return SelectAtPoint
        End Function

        ''' <summary>
        ''' Seleciona as entidades contidas em um ponto (Método de pesquisa com base no banco de dados do desenho. Seleciona em toda a área gráfica mesmo que não visível)
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Point3d">Ponto a ser pesquisado</param>
        ''' <param name="AllowedClassCollection">Classes com os tipos de entidades a serem observadas</param>
        ''' <param name="XDataKeys">Chaves XData a serem procuradas</param>
        ''' <param name="IgnoreIDs">ID´s a serem ignorados</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SelectAtPoint(Transaction As Transaction, Point3d As Point3d, AllowedClassCollection As List(Of System.Type), Optional XDataOperator As eXDataOperator = eXDataOperator.None, Optional XDataKeys As List(Of String) = Nothing, Optional IgnoreIDs As ObjectIdCollection = Nothing) As ObjectIdCollection
            SelectAtPoint = New ObjectIdCollection
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim DBObject As DBObject
            Dim DBObjectCollection As New DBObjectCollection
            Dim AllowedClassCollection2 As New List(Of System.Type)
            Dim Valid As Boolean = False
            Dim Count As Integer = 0
            BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForRead)
            For Each ObjectId As ObjectId In BlockTableRecord
                DBObject = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                If IsNothing(XDataKeys) = False And IsNothing(DBObject.XData) = False Then
                    Select Case XDataOperator
                        Case eXDataOperator._AND
                            For Each XDataKey As String In XDataKeys
                                If DBObject.ContainsXData(XDataKey) = True Then
                                    Count = (Count + 1)
                                End If
                            Next
                            If Count.Equals(XDataKeys.Count) Then
                                Valid = True
                            Else
                                Valid = False
                            End If
                        Case eXDataOperator._OR
                            For Each XDataKey As String In XDataKeys
                                If DBObject.ContainsXData(XDataKey) = True Then
                                    Valid = True
                                    Exit For
                                End If
                            Next
                        Case Else
                            Throw New System.Exception("Parâmetro XDataOperator não informado.")
                    End Select
                Else
                    Valid = True
                End If
                If AllowedClassCollection.Contains(DBObject.GetType) = True And Valid = True And If(IsNothing(IgnoreIDs) = False, If(IgnoreIDs.Contains(DBObject.ObjectId) = True, 0 = 1, 1 = 1), 1 = 1) Then
                    If Engine2.EntityInteration.IsInsidePoint(DBObject, Point3d) = True Then
                        SelectAtPoint.Add(ObjectId)
                    End If
                End If
            Next
            Return SelectAtPoint
        End Function

        ''' <summary>
        ''' Seleciona entidades em um ponto específico (Seleciona apenas na área visível)
        ''' </summary>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="TypedValueCollection">Filtro para seleção de entidades</param>
        ''' <param name="IgnoreIds">Id´s a serem ignorados no processamento</param>
        ''' <param name="Reach">Número que representa o alcance da seleção com base no ponto</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectAtPoint(PointReference As Point3d, TypedValueCollection As List(Of TypedValue), Optional IgnoreIds As ObjectIdCollection = Nothing, Optional SelectAtPointRules As eSelectAtPointRules = eSelectAtPointRules.PassingThroughPoint, Optional Reach As Double = 0.01) As DBObjectCollection
            SelectAtPoint = New DBObjectCollection
            Dim Index As Integer = 0
            Dim EntityInteractionCollection As EntityInteractionCollection
            SelectAtPoint = Engine2.AcadInterface.SelectReach(PointReference, Reach, TypedValueCollection, eSelectReachMode.SelectCrossingPolygon, False, False,, 8)
            EntityInteractionCollection = New EntityInteractionCollection(SelectAtPoint, PointReference, SelectAtPointRules, IgnoreIds)
            EntityInteractionCollection = EntityInteractionCollection.Filter(AddressOf Engine2.AcadInterface.SelectAtPointFilter)
            SelectAtPoint = EntityInteractionCollection.ToDBObjectCollection
            Return SelectAtPoint
        End Function

        ''' <summary>
        ''' Regras para seleção para função SelectAtPoint
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eSelectAtPointRules

            ''' <summary>
            ''' Considera apenas as entidades que comecem ou terminem no ponto
            ''' </summary>
            ''' <remarks></remarks>
            ThatBeginOrEndPoint = 0

            ''' <summary>
            ''' Considerar entidades que passem pelo ponto (indiferente de iniciar ou terminar nele)
            ''' </summary>
            ''' <remarks></remarks>
            PassingThroughPoint = 1

        End Enum

        ''' <summary>
        ''' Filtro de SelectAtPoint
        ''' </summary>
        ''' <param name="EntityInteraction">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function SelectAtPointFilter(EntityInteraction As EntityInteraction) As Boolean
            Dim Curve As Curve
            Dim Points As Point3dCollection
            Dim Entity As Entity
            Dim PointReference As Point3d
            Dim PerpendicularInfor As PerpendicularInfor
            Dim DBObjectCollection As DBObjectCollection
            Dim DBPoint As DBPoint
            Dim SelectAtPointRules As eSelectAtPointRules
            Dim IgnoreIds As ObjectIdCollection
            Entity = EntityInteraction.Entity
            PointReference = EntityInteraction.EntityInteractionCollection.PointReference
            SelectAtPointRules = EntityInteraction.EntityInteractionCollection.SelectAtPointRules
            IgnoreIds = EntityInteraction.EntityInteractionCollection.IgnoreIds
            If Engine2.EntityInteration.IsCurved(Entity) = True Then
                Curve = Entity
                Points = New Point3dCollection
                Points.Add(Curve.StartPoint)
                Points.Add(Curve.EndPoint)
                If Engine2.Geometry.GetClosestPoint(Points, PointReference).DistanceTo(PointReference) <= Tolerance.Global.EqualPoint Then
                    If IgnoreIds.Contains(Entity.ObjectId) = True Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    If SelectAtPointRules = eSelectAtPointRules.PassingThroughPoint Then
                        PerpendicularInfor = Engine2.Geometry.GetPerpendicular(Entity, PointReference)
                        If IsNothing(PerpendicularInfor) = False Then
                            If PerpendicularInfor.VirtualPerpendicular = False Then
                                If PointReference.DistanceTo(PerpendicularInfor.PerpendicularPoint) <= Tolerance.Global.EqualPoint Then
                                    If IsNothing(IgnoreIds) = False Then
                                        If IgnoreIds.Contains(Entity.ObjectId) = True Then
                                            Return False
                                        Else
                                            Return True
                                        End If
                                    Else
                                        Return True
                                    End If
                                Else
                                    Return False
                                End If
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                End If
            Else
                Select Case Entity.GetType.Name
                    Case "BlockReference"
                        DBObjectCollection = Engine2.EntityInteration.GetSubEntitys(Entity, False, PointReference, , , , True)
                        If DBObjectCollection.Count > 0 Then
                            If IsNothing(IgnoreIds) = False Then
                                If IgnoreIds.Contains(Entity.ObjectId) = True Then
                                    Return False
                                Else
                                    Return True
                                End If
                            Else
                                Return True
                            End If
                        Else
                            Return False
                        End If
                    Case "DBPoint"
                        DBPoint = Entity
                        If PointReference.DistanceTo(DBPoint.Position) <= Tolerance.Global.EqualPoint Then
                            If IsNothing(IgnoreIds) = False Then
                                If IgnoreIds.Contains(Entity.ObjectId) = True Then
                                    Return False
                                Else
                                    Return True
                                End If
                            Else
                                Return True
                            End If
                        Else
                            Return False
                        End If
                    Case Else
                        Return False
                End Select
            End If
        End Function

        ''' <summary>
        ''' Classe de agrupamento de entidades para a função SelectAtPoint.
        ''' Permite filtrar a coleção por meio da função SelectAtPointFilter
        ''' </summary>
        ''' <remarks></remarks>
        Public Class EntityInteractionCollection : Inherits List(Of EntityInteraction)

            ''' <summary>
            ''' Modo de seleção
            ''' </summary>
            ''' <remarks></remarks>
            Private _SelectAtPointRules As eSelectAtPointRules

            ''' <summary>
            ''' Ponto de referência
            ''' </summary>
            ''' <remarks></remarks>
            Private _PointReference As Point3d

            ''' <summary>
            ''' Relação de ID´s a serem ignorados
            ''' </summary>
            ''' <remarks></remarks>
            Private _IgnoreIds As ObjectIdCollection

            ''' <summary>
            ''' Modo de seleção
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property SelectAtPointRules As eSelectAtPointRules
                Get
                    Return Me._SelectAtPointRules
                End Get
            End Property

            ''' <summary>
            ''' Retorna o ponto de referência
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property PointReference As Point3d
                Get
                    Return Me._PointReference
                End Get
            End Property

            ''' <summary>
            ''' Relação de ID´s a serem ignorados
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property IgnoreIds As ObjectIdCollection
                Get
                    Return Me._IgnoreIds
                End Get
            End Property

            ''' <summary>
            ''' Retorna a lista de DBObject´s
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function ToDBObjectCollection() As DBObjectCollection
                Dim DBObjectCollection = New DBObjectCollection
                For Index As Integer = 0 To MyBase.Count - 1
                    DBObjectCollection.Add(MyBase.Item(Index).Entity)
                Next
                Return DBObjectCollection
            End Function

            ''' <summary>
            ''' Função de filtro
            ''' </summary>
            ''' <param name="Match"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Filter(Match As System.Predicate(Of EntityInteraction)) As EntityInteractionCollection
                Filter = New EntityInteractionCollection(MyBase.FindAll(Match))
                Return Filter
            End Function

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="DBObjectCollection">DBObjectCollection</param>
            ''' <param name="PointReference">Ponto de referência</param>
            ''' <param name="SelectAtPointRules">Modo de seleção</param>
            ''' <param name="IgnoreIds">Relação de ID´s a serem ignorados</param>
            ''' <remarks></remarks>
            Public Sub New(DBObjectCollection As DBObjectCollection, PointReference As Point3d, SelectAtPointRules As eSelectAtPointRules, IgnoreIds As ObjectIdCollection)
                For Each DbObject As DBObject In DBObjectCollection
                    MyBase.Add(New EntityInteraction(Me, DbObject))
                Next
                Me._PointReference = PointReference
                Me._SelectAtPointRules = SelectAtPointRules
                Me._IgnoreIds = If(IsNothing(IgnoreIds) = True, New ObjectIdCollection, IgnoreIds)
            End Sub

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="EntityCollection">Coleção de entity</param>
            ''' <remarks></remarks>
            Public Sub New(EntityCollection As List(Of EntityInteraction))
                If EntityCollection.Count > 0 Then
                    For Each EntityInteraction As EntityInteraction In EntityCollection
                        MyBase.Add(EntityInteraction)
                    Next
                    Me._PointReference = MyBase.Item(0).EntityInteractionCollection.PointReference
                    Me._SelectAtPointRules = MyBase.Item(0).EntityInteractionCollection.SelectAtPointRules
                    Me._IgnoreIds = MyBase.Item(0).EntityInteractionCollection.IgnoreIds
                End If
            End Sub

        End Class

        ''' <summary>
        ''' Membro de EntityInteractionCollection
        ''' </summary>
        ''' <remarks></remarks>
        Public Class EntityInteraction

            ''' <summary>
            ''' Entidade
            ''' </summary>
            ''' <remarks></remarks>
            Private _Entity As Entity

            ''' <summary>
            ''' Classe mãe
            ''' </summary>
            ''' <remarks></remarks>
            Private _EntityInteractionCollection As EntityInteractionCollection

            ''' <summary>
            ''' Retorna a entidade
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Entity As Entity
                Get
                    Return Me._Entity
                End Get
            End Property

            ''' <summary>
            ''' Retorna a classe mãe
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property EntityInteractionCollection As EntityInteractionCollection
                Get
                    Return Me._EntityInteractionCollection
                End Get
            End Property

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="EntityInteractionCollection">EntityInteractionCollection</param>
            ''' <param name="Entity">Entidade</param>
            ''' <remarks></remarks>
            Public Sub New(EntityInteractionCollection As EntityInteractionCollection, Entity As Entity)
                Me._EntityInteractionCollection = EntityInteractionCollection
                Me._Entity = Entity
            End Sub

        End Class

#End Region

#Region "SelectSequence"

        ''' <summary>
        ''' Seleciona entidades interligadas com base em entidade inicial indicada
        ''' </summary>
        ''' <param name="StartSelection">Entidade inicial para seleção</param>
        ''' <param name="FilterSelectSequenceOptions">Filtro de seleção</param>
        ''' <param name="UseHighlight">Determina se deve destacar os itens localizados na sequencia</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SelectCurveSequence(StartSelection As Entity, FilterSelectSequenceOptions As FilterSelectSequenceOptions, Optional UseHighlight As Boolean = True) As SelectSequenceResult

            'Cria a variável de retorno
            SelectCurveSequence = New SelectSequenceResult

            'Declarações
            Dim ToProcess As New DBObjectCollection
            Dim BaseItem As Entity
            Dim Curve As Curve
            Dim Curve2 As Curve
            Dim TypedValueCollection As New List(Of TypedValue)
            Dim IgnoreIds As New ObjectIdCollection
            Dim Selected As DBObjectCollection
            Dim BasePointType As eAssocPointType
            Dim AssocPointType As eAssocPointType
            Dim AssocPoint As Point3d
            Dim ValidItems As New ArrayList

            '####################################################
            'CRIA A RELAÇÃO DE ITENS VÁLIDOS
            '####################################################
            If FilterSelectSequenceOptions.UseArc = True Then
                ValidItems.Add("Arc")
            End If
            If FilterSelectSequenceOptions.UseLine = True Then
                ValidItems.Add("Line")
            End If
            If FilterSelectSequenceOptions.UsePolyline = True Then
                ValidItems.Add("Polyline")
            End If
            If FilterSelectSequenceOptions.UseSpline = True Then
                ValidItems.Add("Spline")
            End If

            'Verifica o tipo (para tratamento)
            If ValidItems.Contains(StartSelection.GetType.Name) = False Then
                Throw New System.Exception("A entidade representada no parãmetro 'StartSelection' não é do tipo esperado em 'FilterSelectSequenceOptions'")
            End If

            '####################################################
            'INICIA A ANALISE
            '####################################################

            'Inclui o item inicial para o processamento
            ToProcess.Add(StartSelection)

            'Destaca o item inicial
            If UseHighlight = True Then
                Engine2.EntityInteration.HighlightOn(StartSelection)
            End If

            'Processa os itens
            While ToProcess.Count <> 0

                'Obtem o item a ser processado
                BaseItem = ToProcess.Item(0)

                'Inclui o item a ser processado na lista de restrição
                IgnoreIds.Add(BaseItem.ObjectId)

                'Tranfere o item a ser processado para elemento curva
                Curve = BaseItem

                '####################################################
                'SELECIONA ENTIDADES NOS PONTOS INICIAL E FINAL DA CURVA
                '####################################################

                'Monta o filtro de seleção
                With TypedValueCollection
                    .Clear()
                    .Add(New TypedValue(DxfCode.Operator, "<OR"))
                    If FilterSelectSequenceOptions.UseArc = True Then
                        .Add(New TypedValue(DxfCode.Start, "ARC"))
                    End If
                    If FilterSelectSequenceOptions.UseLine = True Then
                        .Add(New TypedValue(DxfCode.Start, "LINE"))
                    End If
                    If FilterSelectSequenceOptions.UsePolyline = True Then
                        .Add(New TypedValue(DxfCode.Start, "LWPOLYLINE"))
                    End If
                    If FilterSelectSequenceOptions.UseSpline = True Then
                        .Add(New TypedValue(DxfCode.Start, "SPLINE"))
                    End If
                    .Add(New TypedValue(DxfCode.Operator, "OR>"))
                    If FilterSelectSequenceOptions.Layers.Count > 0 Then
                        .Add(New TypedValue(DxfCode.Operator, "<OR"))
                        For Each Layer As String In FilterSelectSequenceOptions.Layers
                            .Add(New TypedValue(DxfCode.LayerName, Layer))
                        Next
                        .Add(New TypedValue(DxfCode.Operator, "OR>"))
                    End If
                End With

                'Cria a seleção para o ponto inicial
                Selected = Engine2.AcadInterface.SelectAtPoint(Curve.StartPoint, TypedValueCollection, IgnoreIds, eSelectAtPointRules.ThatBeginOrEndPoint)

                If Selected.Count > 0 Then

                    'Obtem o tipo de ponto de ancoragem na base
                    BasePointType = eAssocPointType.StartPoint

                    'Avalia o retorno
                    For Each AssocItem As DBObject In Selected

                        'Obtem o tipo de ponto de ancoragem na associação
                        Curve2 = AssocItem
                        If Curve.StartPoint.DistanceTo(Curve2.StartPoint) <= Tolerance.Global.EqualPoint Then
                            AssocPointType = eAssocPointType.StartPoint
                            AssocPoint = Curve2.StartPoint
                        ElseIf Curve.StartPoint.DistanceTo(Curve2.EndPoint) <= Tolerance.Global.EqualPoint Then
                            AssocPointType = eAssocPointType.EndPoint
                            AssocPoint = Curve2.EndPoint
                        End If

                        'Adiciona o item na lista de processados
                        SelectCurveSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                        'Adiciona o item associado para análise
                        If ToProcess.Contains(AssocItem) = False Then
                            ToProcess.Add(AssocItem)
                        End If

                    Next

                End If

                '####################################################

                'Cria a seleção para o ponto final
                Selected = Engine2.AcadInterface.SelectAtPoint(Curve.EndPoint, TypedValueCollection, IgnoreIds, eSelectAtPointRules.ThatBeginOrEndPoint)

                If Selected.Count > 0 Then

                    'Obtem o tipo de ponto de ancoragem na base
                    BasePointType = eAssocPointType.EndPoint

                    'Avalia o retorno
                    For Each AssocItem As DBObject In Selected

                        'Obtem o tipo de ponto de ancoragem na associação
                        Curve2 = AssocItem
                        If Curve.EndPoint.DistanceTo(Curve2.StartPoint) <= Tolerance.Global.EqualPoint Then
                            AssocPointType = eAssocPointType.StartPoint
                            AssocPoint = Curve2.StartPoint
                        ElseIf Curve.EndPoint.DistanceTo(Curve2.EndPoint) <= Tolerance.Global.EqualPoint Then
                            AssocPointType = eAssocPointType.EndPoint
                            AssocPoint = Curve2.EndPoint
                        End If

                        'Adiciona o item na lista de processados
                        SelectCurveSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                        'Adiciona o item associado para análise
                        If ToProcess.Contains(AssocItem) = False Then
                            ToProcess.Add(AssocItem)
                        End If

                    Next

                End If

                'Destaca o item
                If UseHighlight = True Then
                    Engine2.EntityInteration.HighlightOn(BaseItem)
                End If

                'Remove o item processado da coleção
                ToProcess.RemoveAt(0)

            End While

            'Informa a quantidade selecionada
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & SelectCurveSequence.Count & " sequencia(s) localizada(s)." & vbLf)

            'Retorno
            Return SelectCurveSequence

        End Function

        ''' <summary>
        ''' Seleciona entidades interligadas com base em entidades limite e entidade inicial para seleção
        ''' </summary>
        ''' <param name="StartSelection">Entidade inicial para seleção</param>
        ''' <param name="Limits">Limites da seleção</param>
        ''' <param name="FilterSelectSequenceOptions">Filtro de seleção</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SelectSequence(StartSelection As Entity, Limits As ObjectIdCollection, FilterSelectSequenceOptions As FilterSelectSequenceOptions) As SelectSequenceResult

            'Cria a variável de retorno
            SelectSequence = New SelectSequenceResult

            'Avalia os parâmetros
            If Limits.Contains(StartSelection.ObjectId) = True Then
                Return SelectSequence
            End If

            'Declarações
            Dim ToProcess As New DBObjectCollection
            Dim BaseItem As Entity
            Dim Curve As Curve
            Dim Curve2 As Curve
            Dim TypedValueCollection As New List(Of TypedValue)
            Dim IgnoreIds As New ObjectIdCollection
            Dim Selected As DBObjectCollection
            Dim ObjectIdCollection As ObjectIdCollection
            Dim BlockCollection As New List(Of BlockReference)
            Dim BlockFilter As List(Of BlockReference)
            Dim BasePointType As eAssocPointType
            Dim AssocPointType As eAssocPointType
            Dim BlockReference As BlockReference
            Dim Intersections As IntersectionEntityCollection
            Dim AssocPoint As Point3d

            '####################################################
            'OBTEM A COLEÇÃO DE BLOCOS DO DESENHO
            '####################################################

            'Monta o filtro de seleção
            With TypedValueCollection
                .Clear()
                .Add(New TypedValue(DxfCode.Operator, "<OR"))
                .Add(New TypedValue(DxfCode.Start, "INSERT"))
                .Add(New TypedValue(DxfCode.Operator, "OR>"))
                If FilterSelectSequenceOptions.Layers.Count > 0 Then
                    .Add(New TypedValue(DxfCode.Operator, "<OR"))
                    For Each Layer As String In FilterSelectSequenceOptions.Layers
                        .Add(New TypedValue(DxfCode.LayerName, Layer))
                    Next
                    .Add(New TypedValue(DxfCode.Operator, "OR>"))
                End If
            End With

            'Seleciona todos os blocos
            ObjectIdCollection = Engine2.AcadInterface.SelectAll(TypedValueCollection)

            'Converte a seleção para coleção compatível com filtro
            For Each ObjectId As ObjectId In ObjectIdCollection
                BlockCollection.Add(Engine2.ConvertObject.ObjectIDToEntity(ObjectId))
            Next

            '####################################################
            'INICIA A ANALISE
            '####################################################

            'Inclui o item inicial para o processamento
            ToProcess.Add(StartSelection)

            'Destaca o item inicial
            Engine2.EntityInteration.HighlightOn(StartSelection)

            'Processa os itens
            While ToProcess.Count <> 0

                'Obtem o item a ser processado
                BaseItem = ToProcess.Item(0)

                If Limits.Contains(BaseItem.ObjectId) = False Then

                    'Inclui o item a ser processado na lista de restrição
                    IgnoreIds.Add(BaseItem.ObjectId)

                    'Verifica o tipo (para tratamento)
                    Select Case BaseItem.GetType.Name
                        Case "Line", "Arc", "Spline", "Polyline"

                            'Tranfere o item a ser processado para elemento curva
                            Curve = BaseItem

                            '####################################################
                            'SELECIONA ENTIDADES NOS PONTOS INICIAL E FINAL DA CURVA
                            '####################################################

                            'Monta o filtro de seleção
                            With TypedValueCollection
                                .Clear()
                                .Add(New TypedValue(DxfCode.Operator, "<OR"))
                                If FilterSelectSequenceOptions.UseArc = True Then
                                    .Add(New TypedValue(DxfCode.Start, "ARC"))
                                End If
                                If FilterSelectSequenceOptions.UseLine = True Then
                                    .Add(New TypedValue(DxfCode.Start, "LINE"))
                                End If
                                If FilterSelectSequenceOptions.UsePolyline = True Then
                                    .Add(New TypedValue(DxfCode.Start, "LWPOLYLINE"))
                                End If
                                If FilterSelectSequenceOptions.UseSpline = True Then
                                    .Add(New TypedValue(DxfCode.Start, "SPLINE"))
                                End If
                                .Add(New TypedValue(DxfCode.Start, "INSERT"))
                                .Add(New TypedValue(DxfCode.Operator, "OR>"))
                                If FilterSelectSequenceOptions.Layers.Count > 0 Then
                                    .Add(New TypedValue(DxfCode.Operator, "<OR"))
                                    For Each Layer As String In FilterSelectSequenceOptions.Layers
                                        .Add(New TypedValue(DxfCode.LayerName, Layer))
                                    Next
                                    .Add(New TypedValue(DxfCode.Operator, "OR>"))
                                End If
                            End With

                            'Cria a seleção para o ponto inicial
                            Selected = Engine2.AcadInterface.SelectAtPoint(Curve.StartPoint, TypedValueCollection, IgnoreIds, AcadInterface.eSelectAtPointRules.PassingThroughPoint)

                            'Obtem o tipo de ponto de ancoragem na base
                            BasePointType = eAssocPointType.StartPoint

                            'Avalia o retorno
                            For Each AssocItem As DBObject In Selected

                                'Obtem o tipo de ponto de ancoragem na associação
                                Select Case AssocItem.GetType.Name
                                    Case "Line", "Arc", "Spline", "Polyline"
                                        Curve2 = AssocItem
                                        If Curve.StartPoint.DistanceTo(Curve2.StartPoint) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.StartPoint
                                            AssocPoint = Curve2.StartPoint
                                        ElseIf Curve.StartPoint.DistanceTo(Curve2.EndPoint) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.EndPoint
                                            AssocPoint = Curve2.EndPoint
                                        End If
                                    Case "BlockReference"
                                        BlockReference = AssocItem
                                        If Curve.StartPoint.DistanceTo(BlockReference.Position) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.PositionAndStartPoint
                                            AssocPoint = BlockReference.Position
                                        Else
                                            AssocPointType = eAssocPointType.OtherAndStartPoint
                                            AssocPoint = Curve.StartPoint
                                        End If
                                End Select

                                'Adiciona o item na lista de processados
                                SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                'Adiciona o item associado para análise
                                If ToProcess.Contains(AssocItem) = False Then
                                    ToProcess.Add(AssocItem)
                                End If

                            Next

                            '####################################################

                            'Cria a seleção para o ponto final
                            Selected = Engine2.AcadInterface.SelectAtPoint(Curve.EndPoint, TypedValueCollection, IgnoreIds, AcadInterface.eSelectAtPointRules.PassingThroughPoint)

                            'Obtem o tipo de ponto de ancoragem na base
                            BasePointType = eAssocPointType.EndPoint

                            'Avalia o retorno
                            For Each AssocItem As DBObject In Selected

                                'Obtem o tipo de ponto de ancoragem na associação
                                Select Case AssocItem.GetType.Name
                                    Case "Line", "Arc", "Spline", "Polyline"
                                        Curve2 = AssocItem
                                        If Curve.EndPoint.DistanceTo(Curve2.StartPoint) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.StartPoint
                                            AssocPoint = Curve2.StartPoint
                                        ElseIf Curve.EndPoint.DistanceTo(Curve2.EndPoint) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.EndPoint
                                            AssocPoint = Curve2.EndPoint
                                        End If
                                    Case "BlockReference"
                                        BlockReference = AssocItem
                                        If Curve.EndPoint.DistanceTo(BlockReference.Position) <= Tolerance.Global.EqualPoint Then
                                            AssocPointType = eAssocPointType.PositionAndEndpoint
                                            AssocPoint = BlockReference.Position
                                        Else
                                            AssocPointType = eAssocPointType.OtherAndEndPoint
                                            AssocPoint = Curve.EndPoint
                                        End If
                                End Select

                                'Adiciona o item na lista de processados
                                SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                'Adiciona o item associado para análise
                                If ToProcess.Contains(AssocItem) = False Then
                                    ToProcess.Add(AssocItem)
                                End If

                            Next

                            '####################################################
                            'SELECIONA BLOCOS NOS PONTOS INICIAL E FINAL DA CURVA (COM BASE NO PONTO DE INSERÇÃO)
                            '####################################################

                            'Cria a seleção para ponto inicial
                            BlockFilter = BlockCollection.FindAll(Function(X As BlockReference) X.Position.DistanceTo(Curve.StartPoint) <= Tolerance.Global.EqualPoint)

                            'Obtem o tipo de ponto de ancoragem na base
                            BasePointType = eAssocPointType.StartPoint

                            'Obtem o tipo de ponto de ancoragem na associação
                            AssocPointType = eAssocPointType.PositionAndStartPoint

                            'Avalia o retorno
                            For Each AssocItem As BlockReference In BlockFilter

                                'Obtem o ponto
                                AssocPoint = AssocItem.Position

                                'Adiciona o item na lista de processados
                                SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                'Adiciona o item associado para análise
                                If ToProcess.Contains(AssocItem) = False Then
                                    ToProcess.Add(AssocItem)
                                End If

                            Next

                            '####################################################

                            'Cria a seleção para ponto final
                            BlockFilter = BlockCollection.FindAll(Function(X As BlockReference) X.Position.DistanceTo(Curve.EndPoint) <= Tolerance.Global.EqualPoint)

                            'Obtem o tipo de ponto de ancoragem na base
                            BasePointType = eAssocPointType.EndPoint

                            'Obtem o tipo de ponto de ancoragem na associação
                            AssocPointType = eAssocPointType.PositionAndEndpoint

                            'Avalia o retorno
                            For Each AssocItem As BlockReference In BlockFilter

                                'Obtem o ponto
                                AssocPoint = AssocItem.Position

                                'Adiciona o item na lista de processados
                                SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                'Adiciona o item associado para análise
                                If ToProcess.Contains(AssocItem) = False Then
                                    ToProcess.Add(AssocItem)
                                End If

                            Next

                        Case "BlockReference"

                            'Tranfere o item a ser processado para elemento BlockReference
                            BlockReference = BaseItem

                            '####################################################
                            'SELECIONA ENTIDADES NA POSIÇÃO DO BLOCO (PONTO DE INSERÇÃO)
                            '####################################################

                            'Monta o filtro de seleção
                            With TypedValueCollection
                                .Clear()
                                .Add(New TypedValue(DxfCode.Operator, "<OR"))
                                If FilterSelectSequenceOptions.UseArc = True Then
                                    .Add(New TypedValue(DxfCode.Start, "ARC"))
                                End If
                                If FilterSelectSequenceOptions.UseLine = True Then
                                    .Add(New TypedValue(DxfCode.Start, "LINE"))
                                End If
                                If FilterSelectSequenceOptions.UsePolyline = True Then
                                    .Add(New TypedValue(DxfCode.Start, "LWPOLYLINE"))
                                End If
                                If FilterSelectSequenceOptions.UseSpline = True Then
                                    .Add(New TypedValue(DxfCode.Start, "SPLINE"))
                                End If
                                .Add(New TypedValue(DxfCode.Operator, "OR>"))
                                If FilterSelectSequenceOptions.Layers.Count > 0 Then
                                    .Add(New TypedValue(DxfCode.Operator, "<OR"))
                                    For Each Layer As String In FilterSelectSequenceOptions.Layers
                                        .Add(New TypedValue(DxfCode.LayerName, Layer))
                                    Next
                                    .Add(New TypedValue(DxfCode.Operator, "OR>"))
                                End If
                            End With

                            'Cria a seleção para posição do bloco
                            Selected = Engine2.AcadInterface.SelectAtPoint(BlockReference.Position, TypedValueCollection, IgnoreIds, AcadInterface.eSelectAtPointRules.ThatBeginOrEndPoint)

                            'Avalia o retorno
                            For Each AssocItem As DBObject In Selected

                                'Obtem a curva para comparação
                                Curve2 = AssocItem

                                If BlockReference.Position.DistanceTo(Curve2.StartPoint) <= Tolerance.Global.EqualPoint Then

                                    'Obtem o tipo de ponto de ancoragem na base
                                    BasePointType = eAssocPointType.PositionAndStartPoint

                                    'Obtem o tipo de ponto de ancoragem na associação
                                    AssocPointType = eAssocPointType.StartPoint

                                    'Obtem o ponto
                                    AssocPoint = Curve2.StartPoint

                                    'Adiciona o item na lista de processados
                                    SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                    'Adiciona o item associado para análise
                                    If ToProcess.Contains(AssocItem) = False Then
                                        ToProcess.Add(AssocItem)
                                    End If

                                ElseIf BlockReference.Position.DistanceTo(Curve2.EndPoint) <= Tolerance.Global.EqualPoint Then

                                    'Obtem o tipo de ponto de ancoragem na base
                                    BasePointType = eAssocPointType.PositionAndEndpoint

                                    'Obtem o tipo de ponto de ancoragem na associação
                                    AssocPointType = eAssocPointType.EndPoint

                                    'Obtem o ponto
                                    AssocPoint = Curve2.EndPoint

                                    'Adiciona o item na lista de processados
                                    SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                    'Adiciona o item associado para análise
                                    If ToProcess.Contains(AssocItem) = False Then
                                        ToProcess.Add(AssocItem)
                                    End If

                                End If

                            Next

                            'Obtem as sub entidades do bloco
                            Selected = Engine2.EntityInteration.GetSubEntitys(BlockReference, True, , New ArrayList({"Line", "Arc", "Spline", "Polyline", "Circle", "Ellipse"}))

                            '####################################################
                            'SELECIONA ENTIDADES POSICIONADOS EM PARTES INTERNAS DO BLOCO
                            '####################################################

                            'Avalia o retorno
                            For Each AssocItem As Entity In Selected

                                'Obtem as intersecções
                                Intersections = Engine2.EntityInteration.GetEntityIntersections(AssocItem, Intersect.OnBothOperands, TypedValueCollection, IgnoreIds)

                                'Avalia o retorno
                                For Each IntersectionEntity As IntersectionEntity In Intersections

                                    'Tranfere o item a ser processado para elemento curva
                                    Curve = IntersectionEntity.Entity

                                    'Obtem o item associado
                                    AssocItem = IntersectionEntity.Entity

                                    'Avalia a entidade começa ou termina no ponto de intersecção 
                                    For Each Point3d As Point3d In IntersectionEntity.IntersectionPoints

                                        If Curve.StartPoint.DistanceTo(Point3d) <= Tolerance.Global.EqualPoint Then

                                            'Obtem o tipo de ponto de ancoragem na base
                                            BasePointType = eAssocPointType.OtherAndStartPoint

                                            'Obtem o ponto
                                            AssocPoint = Curve.StartPoint

                                            'Obtem o tipo de ponto de ancoragem na associação
                                            AssocPointType = eAssocPointType.StartPoint

                                            'Adiciona o item na lista de processados
                                            SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                            'Adiciona o item associado para análise
                                            If ToProcess.Contains(AssocItem) = False Then
                                                ToProcess.Add(AssocItem)
                                            End If

                                        ElseIf Curve.EndPoint.DistanceTo(Point3d) <= Tolerance.Global.EqualPoint Then

                                            'Obtem o tipo de ponto de ancoragem na base
                                            BasePointType = eAssocPointType.OtherAndEndPoint

                                            'Obtem o ponto
                                            AssocPoint = Curve.EndPoint

                                            'Obtem o tipo de ponto de ancoragem na associação
                                            AssocPointType = eAssocPointType.EndPoint

                                            'Adiciona o item na lista de processados
                                            SelectSequence.Add(BaseItem, AssocItem, AssocPoint, BasePointType, AssocPointType)

                                            'Adiciona o item associado para análise
                                            If ToProcess.Contains(AssocItem) = False Then
                                                ToProcess.Add(AssocItem)
                                            End If

                                        End If

                                    Next

                                Next

                            Next

                    End Select

                End If

                'Destaca o item
                Engine2.EntityInteration.HighlightOn(BaseItem)

                'Remove o item processado da coleção
                ToProcess.RemoveAt(0)

            End While

            'Informa a quantidade selecionada
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & SelectSequence.Count & " sequencia(s) localizada(s)." & vbLf)

            'Retorno
            Return SelectSequence

        End Function

        ''' <summary>
        ''' Retorno de SelectSequence
        ''' </summary>
        ''' <remarks></remarks>
        Public NotInheritable Class SelectSequenceResult : Inherits List(Of SelectSequenceItem)

            ''' <summary>
            ''' Adiciona item
            ''' </summary>
            ''' <param name="BaseEntity">Entidade base da seleção</param>
            ''' <param name="AssocEntity">Entidade associada</param>
            ''' <param name="AnchorPoint">Ponto de ancoragem</param>
            ''' <param name="BasePointType">Identificação do ponto associado na entidade base</param>
            ''' <param name="AssocPointType">Identificação do ponto associado na entidade associada</param>
            ''' <remarks></remarks>
            Public Overloads Sub Add(BaseEntity As Entity, AssocEntity As Entity, AnchorPoint As Point3d, BasePointType As eAssocPointType, AssocPointType As eAssocPointType)
                If IsNothing(MyBase.Find(Function(X As SelectSequenceItem) X.BaseEntity = BaseEntity And X.AssocEntity = AssocEntity)) = True And IsNothing(MyBase.Find(Function(X As SelectSequenceItem) X.BaseEntity = AssocEntity And X.AssocEntity = BaseEntity)) = True Then
                    MyBase.Add(New SelectSequenceItem(Me, BaseEntity, AssocEntity, AnchorPoint, BasePointType, AssocPointType))
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "-> Sequencia localizada entre '" & BaseEntity.GetType.Name & " (" & BaseEntity.Handle.ToString & ")' e '" & AssocEntity.GetType.Name & " (" & AssocEntity.Handle.ToString & ")'." & vbLf)
                    System.Windows.Forms.Application.DoEvents()
                End If
            End Sub

            ''' <summary>
            ''' Verifica e existência de um item
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function ContainsBaseEntity(BaseEntity As Entity) As Boolean
                If IsNothing(MyBase.Find(Function(X As SelectSequenceItem) X.BaseEntity.Equals(BaseEntity))) = True Then
                    Return False
                Else
                    Return True
                End If
            End Function

            ''' <summary>
            ''' Verifica e existência de um item
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function ContainsAssocEntity(BaseEntity As Entity) As Boolean
                If IsNothing(MyBase.Find(Function(X As SelectSequenceItem) X.AssocEntity.Equals(BaseEntity))) = True Then
                    Return False
                Else
                    Return True
                End If
            End Function

        End Class

        ''' <summary>
        ''' Item de SelectSequenceResult
        ''' </summary>
        ''' <remarks></remarks>
        Public NotInheritable Class SelectSequenceItem

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="SelectSequenceResult">Classe mãe</param>
            ''' <param name="BaseEntity">Entidade base da seleção</param>
            ''' <param name="AssocEntity">Entidade associada</param>
            ''' <param name="AnchorPoint">Ponto de ancoragem</param>
            ''' <param name="BasePointType">Identificação do ponto associado na entidade base</param>
            ''' <param name="AssocPointType">Identificação do ponto associado na entidade associada</param>
            ''' <remarks></remarks>
            Public Sub New(SelectSequenceResult As SelectSequenceResult, BaseEntity As Entity, AssocEntity As Entity, AnchorPoint As Point3d, BasePointType As eAssocPointType, AssocPointType As eAssocPointType)
                Me._SelectSequenceResult = SelectSequenceResult
                Me._BaseEntity = BaseEntity
                Me._AssocEntity = AssocEntity
                Me._AssocPoint = AnchorPoint
                Me._BasePointType = BasePointType
                Me._AssocPointType = AssocPointType
            End Sub

#Region "ENUMERADORES"

            ''' <summary>
            ''' Determina qual o ponto associado para 'AssocEntity'
            ''' </summary>
            ''' <remarks></remarks>
            Public Enum eAssocPointType

                ''' <summary>
                ''' Ponto inicial
                ''' </summary>
                ''' <remarks></remarks>
                StartPoint = 0

                ''' <summary>
                ''' Ponto final
                ''' </summary>
                ''' <remarks></remarks>
                EndPoint = 1

                ''' <summary>
                ''' Outro ponto e ponto inicial
                ''' </summary>
                ''' <remarks></remarks>
                OtherAndStartPoint = 2

                ''' <summary>
                ''' Outro ponto e ponto final
                ''' </summary>
                ''' <remarks></remarks>
                OtherAndEndPoint = 3

                ''' <summary>
                ''' Posição do bloco e ponto inicial
                ''' </summary>
                ''' <remarks></remarks>
                PositionAndStartPoint = 4

                ''' <summary>
                ''' Posição do bloco e ponto final
                ''' </summary>
                ''' <remarks></remarks>
                PositionAndEndpoint = 5

            End Enum

#End Region

#Region "VARIÁVEIS"

            ''' <summary>
            ''' Classe mãe
            ''' </summary>
            ''' <remarks></remarks>
            Private _SelectSequenceResult As SelectSequenceResult

            ''' <summary>
            ''' Entidade base da seleção
            ''' </summary>
            ''' <remarks></remarks>
            Private _BaseEntity As Entity

            ''' <summary>
            ''' Entidade associada
            ''' </summary>
            ''' <remarks></remarks>
            Private _AssocEntity As Entity

            ''' <summary>
            ''' Ponto de ancoragem
            ''' </summary>
            ''' <remarks></remarks>
            Private _AssocPoint As Point3d

            ''' <summary>
            ''' Identificação do ponto associado na entidade base
            ''' </summary>
            ''' <remarks></remarks>
            Private _BasePointType As eAssocPointType

            ''' <summary>
            ''' Identificação do ponto associado na entidade associada
            ''' </summary>
            ''' <remarks></remarks>
            Private _AssocPointType As eAssocPointType

#End Region

#Region "PROPRIEDADES"

            ''' <summary>
            ''' Classe mãe
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property SelectSequenceResult As SelectSequenceResult
                Get
                    Return Me._SelectSequenceResult
                End Get
            End Property

            ''' <summary>
            ''' Entidade base da seleção
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property BaseEntity As Entity
                Get
                    Return Me._BaseEntity
                End Get
                Set(value As Entity)
                    Me.UpdateData(Me._BaseEntity, value)
                    Me._BaseEntity = value
                End Set
            End Property

            ''' <summary>
            ''' Entidade associada
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property AssocEntity As Entity
                Get
                    Return Me._AssocEntity
                End Get
                Set(value As Entity)
                    Me.UpdateData(Me._AssocEntity, value)
                    Me._AssocEntity = value
                End Set
            End Property

            ''' <summary>
            ''' Ponto de ancoragem
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property AssocPoint As Point3d
                Get
                    Return Me._AssocPoint
                End Get
            End Property

            ''' <summary>
            ''' Identificação do ponto associado na entidade base
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property BasePointType As eAssocPointType
                Get
                    Return Me._BasePointType
                End Get
            End Property

            ''' <summary>
            ''' Identificação do ponto associado na entidade associada
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property AssocPointType As eAssocPointType
                Get
                    Return Me._AssocPointType
                End Get
            End Property

#End Region

#Region "MÉTODOS"

            ''' <summary>
            ''' Atualiza a coleção
            ''' </summary>
            ''' <param name="ActualEntity">Entidade atual</param>
            ''' <param name="NewEntity">Nova entidade</param>
            ''' <remarks></remarks>
            Private Sub UpdateData(ActualEntity As Entity, NewEntity As Entity)
                Dim Filter As List(Of SelectSequenceItem)
                Filter = Me._SelectSequenceResult.FindAll(Function(X As SelectSequenceItem) ((X.BaseEntity.Equals(ActualEntity) = True Or X.AssocEntity.Equals(ActualEntity) = True) And X.Equals(Me) = False))
                For Each SelectSequenceItem As SelectSequenceItem In Filter
                    With SelectSequenceItem
                        If .BaseEntity = ActualEntity Then
                            ._BaseEntity = NewEntity
                        ElseIf .AssocEntity = ActualEntity Then
                            ._AssocEntity = ActualEntity
                        End If
                    End With
                Next
            End Sub

#End Region

        End Class

        ''' <summary>
        ''' Opções de filtro para seleção de sequências
        ''' </summary>
        ''' <remarks></remarks>
        Public NotInheritable Class FilterSelectSequenceOptions

            ''' <summary>
            ''' Contrutor
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New()
                Me._UseArc = True
                Me._UseLine = True
                Me._UsePolyline = True
                Me._UseSpline = True
                Me._Layers = New List(Of String)
            End Sub

            ''' <summary>
            ''' Contrutor
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub New(UseArc As Boolean, UseLine As Boolean, UsePolyline As Boolean, UseSpline As Boolean, Layers As List(Of String))
                Me._UseArc = UseArc
                Me._UseLine = UseLine
                Me._UsePolyline = UsePolyline
                Me._UseSpline = UseSpline
                Me._Layers = Layers
            End Sub

#Region "VARIÁVEIS"

            ''' <summary>
            ''' Usar Polyline
            ''' </summary>
            ''' <remarks></remarks>
            Private _UsePolyline As Boolean

            ''' <summary>
            ''' Usar Linha
            ''' </summary>
            ''' <remarks></remarks>
            Private _UseLine As Boolean

            ''' <summary>
            ''' Usar Arco
            ''' </summary>
            ''' <remarks></remarks>
            Private _UseArc As Boolean

            ''' <summary>
            ''' Usar Spline
            ''' </summary>
            ''' <remarks></remarks>
            Private _UseSpline As Boolean

            ''' <summary>
            ''' Coleção de camadas
            ''' </summary>
            ''' <remarks></remarks>
            ''' 
            Private _Layers As List(Of String)
#End Region

#Region "PROPRIEDADES"

            ''' <summary>
            ''' Usar Polyline
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property UsePolyline As Boolean
                Get
                    Return Me._UsePolyline
                End Get
                Set(value As Boolean)
                    Me._UsePolyline = value
                End Set
            End Property

            ''' <summary>
            ''' Usar Linha
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property UseLine As Boolean
                Get
                    Return Me._UseLine
                End Get
                Set(value As Boolean)
                    Me._UseLine = value
                End Set
            End Property

            ''' <summary>
            ''' Usar Arco
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property UseArc As Boolean
                Get
                    Return Me._UseArc
                End Get
                Set(value As Boolean)
                    Me._UseArc = value
                End Set
            End Property

            ''' <summary>
            ''' Usar Spline
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property UseSpline As Boolean
                Get
                    Return Me._UseSpline
                End Get
                Set(value As Boolean)
                    Me._UseSpline = value
                End Set
            End Property

            ''' <summary>
            ''' Coleção de camadas
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property Layers As List(Of String)
                Get
                    Return Me._Layers
                End Get
                Set(value As List(Of String))
                    Me._Layers = value
                End Set
            End Property

#End Region

        End Class

#End Region

    End Class

End Namespace

