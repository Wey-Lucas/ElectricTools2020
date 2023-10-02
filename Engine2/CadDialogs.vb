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
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows

Namespace Engine2

    ''' <summary>
    ''' Acesso as janelas do AutoCAD
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CadDialogs

        ''' <summary>
        ''' Abre a caixa de diálogo para salvar arquivos
        ''' </summary>
        ''' <param name="Title">Título</param>
        ''' <param name="Filter">Filtro</param>
        ''' <param name="InitialDirectory">Diretório inicial</param>
        ''' <param name="DefaultExt">Extensão padrão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SaveFile(Title As String, Filter As String, InitialDirectory As String, DefaultExt As String) As Object
            Dim OpenFileDialog As New System.Windows.Forms.SaveFileDialog
            With OpenFileDialog
                .Title = Title
                .FileName = ""
                .DefaultExt = DefaultExt
                .AddExtension = True
                .Filter = Filter
                If InitialDirectory.Trim <> "" Then
                    .InitialDirectory = InitialDirectory
                End If
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Return .FileName
                Else
                    Return Nothing
                End If
            End With
        End Function

        ''' <summary>
        ''' Abre a caixa de diálogo para seleção de arquivos
        ''' </summary>
        ''' <param name="Title">Título</param>
        ''' <param name="Filter">Filtro</param>
        ''' <param name="InitialDirectory">Diretório inicial</param>
        ''' <param name="DefaultExt">Extensão padrão</param>
        ''' <param name="Multiselect">Determina se o sistema permite múltiplas seleções</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OpenFile(Title As String, Filter As String, InitialDirectory As String, DefaultExt As String, Optional Multiselect As Boolean = False) As Object
            Dim OpenFileDialog As New System.Windows.Forms.OpenFileDialog
            With OpenFileDialog
                .Title = Title
                .FileName = ""
                .DefaultExt = DefaultExt
                .AddExtension = True
                .Filter = Filter
                .Multiselect = Multiselect
                If InitialDirectory.Trim <> "" Then
                    .InitialDirectory = InitialDirectory
                End If
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    If Multiselect = True Then
                        Return .FileNames
                    Else
                        Return .FileName
                    End If
                Else
                    Return Nothing
                End If
            End With
        End Function

        ''' <summary>
        ''' Paleta de cores
        ''' </summary>
        ''' <param name="Color">Cor padrão</param>
        ''' <param name="TrueColorTab">TrueColorTab</param>
        ''' <param name="ColorBookTab">ColorBookTab</param>
        ''' <param name="ACITab">ACITab</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ShowColorDialog(Color As Autodesk.AutoCAD.Colors.Color, Optional TrueColorTab As Boolean = True, Optional ColorBookTab As Boolean = True, Optional ACITab As Boolean = True) As Autodesk.AutoCAD.Colors.Color
            Dim ColorDialog As New ColorDialog()
            With ColorDialog
                .Color = Color
                Dim TabColor As Integer = 0
                If ACITab = True Then
                    TabColor = TabColor + ColorDialog.ColorTabs.ACITab
                End If
                If TrueColorTab = True Then
                    TabColor = TabColor + ColorDialog.ColorTabs.TrueColorTab
                End If
                If ColorBookTab = True Then
                    TabColor = TabColor + ColorDialog.ColorTabs.ColorBookTab
                End If
                If TabColor = 0 Then
                    Throw New System.Exception("Nenhuma paleta de cor foi informada")
                End If
                .SetDialogTabs(TabColor)
                .IncludeByBlockByLayer = False
            End With
            Dim dr As System.Windows.Forms.DialogResult = ColorDialog.ShowDialog()
            If dr = System.Windows.Forms.DialogResult.OK Then
                Return ColorDialog.Color
            Else
                Return Nothing
            End If

        End Function

        ''' <summary>
        ''' Caixa para espessura de linha
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ShowLineWeightDialog() As Autodesk.AutoCAD.DatabaseServices.LineWeight
            Dim LineWeightDialog As New LineWeightDialog()
            Dim DialogResult As System.Windows.Forms.DialogResult = LineWeightDialog.ShowDialog()
            If DialogResult = System.Windows.Forms.DialogResult.OK Then
                Return LineWeightDialog.LineWeight
            Else
                Return Nothing
            End If
        End Function

    End Class

End Namespace


