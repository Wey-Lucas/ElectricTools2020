
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows
Imports System.ComponentModel

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Commands.ELECTRICTOOLS))>
Namespace Commands

    Public Class ELECTRICTOOLS

        'Declarações
        Private Shared WithEvents _PaletteSetControl As Engine2.MyPaletteSet
        Private Shared _CurrentPalette As Palette
        Private Shared _ctrElectricTools As ctrElectricTools

        ''' <summary>
        ''' Retorna ctrDTSIMP
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property ctrElectricTools As ctrElectricTools
            Get
                Return _ctrElectricTools
            End Get
        End Property

        <CommandMethod("ELECTRICTOOLS", CommandFlags.Modal)>
        Public Shared Sub ELECTRICTOOLS()

            If Support.Lic.GetLic = False Then
                Exit Sub
            End If

            Support.Lic.Add(Support.Lic.eArchitecture.AutoCAD, My.Application.Info.AssemblyName, "ELECTRICALTOOLS")

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database

            'Avalia se está aberta
            If IsActive() = False Then

                'Carrega a janela
                LoadElectricTools()

                'Mensagem
                Editor.WriteMessage(vbLf & "Paleta ElectricTools aberta..." & vbLf)

            Else

                'Descarrega a janela
                UnLoadElectricTools()

                'Mensagem
                Editor.WriteMessage(vbLf & "Paleta ElectricTools fechada..." & vbLf)

            End If

        End Sub
        <CommandMethod("ELT", CommandFlags.Modal)>
        Public Shared Sub ELT()
            ELECTRICTOOLS()
        End Sub

        ''' <summary>
        ''' Determina se a paleta está ativa
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function IsActive() As Boolean

            'Avalia 
            If IsNothing(_ctrElectricTools) = True Then

                'Retorno
                Return False

            Else

                'Avalia
                If _CurrentPalette.PaletteSet.Visible = True Then

                    'Retorno
                    Return True

                Else

                    'Retorno
                    Return False

                End If
            End If

        End Function

        ''' <summary>
        ''' Carrega a paleta ElectricTools
        ''' </summary>
        Public Shared Sub LoadElectricTools()
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            If IsNothing(_PaletteSetControl) = True Then
                _PaletteSetControl = New Engine2.MyPaletteSet("ELECTRICTOOLS", "ELECTRICTOOLS", New Guid("{DCBFEC73-9FBB-Caa2-8CCC-1068E94A2BFA}"))
                _PaletteSetControl.Icon = My.Resources.LUZ
            End If
            With _PaletteSetControl
                .KeepFocus = False
                If .Visible = False Then
                    While .Count <> 0
                        .Remove(0)
                    End While
                    If .Count = 0 Then
                        'ELECTRICTOOLS
                        _ctrElectricTools = New ctrElectricTools
                        _ctrElectricTools.Visible = True
                        .Add("Tools", _ctrElectricTools)
                        _ctrElectricTools.Dock = Windows.Forms.DockStyle.Fill
                        .DockEnabled = DockSides.Left + DockSides.Right
                        .Visible = True
                        .Style = PaletteSetStyles.ShowCloseButton + PaletteSetStyles.ShowAutoHideButton + PaletteSetStyles.ShowPropertiesMenu + PaletteSetStyles.ShowTabForSingle
                    End If
                    .Activate(0)
                End If
            End With
        End Sub

        ''' <summary>
        ''' Controla o acesso
        ''' </summary>
        Public Shared Sub AccessUpdate()
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If IsNothing(_PaletteSetControl) = False Then
                With _PaletteSetControl
                    .KeepFocus = False
                    If IsNothing(Commands.ELECTRICTOOLS.ctrElectricTools) = False AndAlso Commands.ELECTRICTOOLS.ctrElectricTools.Visible = True Then
                        If IsNothing(Document) = True Then
                            Commands.ELECTRICTOOLS.ctrElectricTools.Enabled = False
                        Else
                            Commands.ELECTRICTOOLS.ctrElectricTools.Enabled = True
                        End If
                    End If
                End With
            End If
        End Sub

        ''' <summary>
        ''' Descarrega a paleta ElectricTools
        ''' </summary>
        Public Shared Sub UnLoadElectricTools()
            If IsNothing(_PaletteSetControl) = False Then
                With _PaletteSetControl
                    .KeepFocus = False
                    If .Visible = True Then
                        While .Count <> 0
                            .Remove(0)
                        End While
                        .Visible = False
                    End If
                End With
            End If
            If IsNothing(_ctrElectricTools) = False Then
                _ctrElectricTools.Dispose()
                _ctrElectricTools = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Retorna PaletteSetControl
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property PaletteSetControl As Engine2.MyPaletteSet
            Get
                Return _PaletteSetControl
            End Get
        End Property

        ''' <summary>
        ''' Retorna a paleta corrente
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property CurrentPalette As Palette
            Get
                Return _CurrentPalette
            End Get
        End Property

        ''' <summary>
        ''' Obtem a paleta corrente
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Shared Sub _PaletteSetControl_PaletteActivated(sender As Object, e As PaletteActivatedEventArgs) Handles _PaletteSetControl.PaletteActivated
            If IsNothing(_PaletteSetControl) = False Then
                If IsNothing(e.Activated) = False Then
                    Select Case e.Activated.Name
                        Case "Tools"
                            If IsNothing(_ctrElectricTools) = False Then
                                Exit Select
                            End If
                    End Select
                    _CurrentPalette = e.Activated
                End If
            End If
        End Sub

    End Class

End Namespace