'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Drawing

Namespace Engine2

    ''' <summary>
    ''' Classe para conversão de unidades métricas
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ConvertionUnits

        ''' <summary>
        ''' Enumerador para os vetores de DPI
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum DpiPos

            ''' <summary>
            ''' DPI em X
            ''' </summary>
            ''' <remarks></remarks>
            X = 0

            ''' <summary>
            ''' DPI em Y
            ''' </summary>
            ''' <remarks></remarks>
            Y = 1

        End Enum

        ''' <summary>
        ''' Converte Milímetro para Pixel
        ''' </summary>
        ''' <param name="Value">Valor em milímetros</param>
        ''' <param name="DpiPos">Posição do vetor</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MmToPx(ByVal Value As Single, ByVal DpiPos As DpiPos) As Single
            Try
                Dim MyGr As Graphics = Graphics.FromImage(New Bitmap(1, 1, Imaging.PixelFormat.Format32bppRgb))
                Select Case DpiPos
                    Case DpiPos.X
                        Return (((MyGr.DpiX / 2.54) * Value) / 10)
                    Case DpiPos.Y
                        Return (((MyGr.DpiY / 2.54) * Value) / 10)
                    Case Else
                        Return 0
                End Select
            Catch ex As System.Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Converte Milimetros para pontos
        ''' </summary>
        ''' <param name="Value">Valor em milímetros</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MmToPt(ByVal Value As Integer) As Object
            Try
                Return ((Value * 56.7) / 20)
            Catch ex As System.Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Converte Milímetro para Twip
        ''' </summary>
        ''' <param name="Value">Valor em milímetros</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MmToTw(ByVal Value As Single) As Single
            Try
                Return (56.7 / Value)
            Catch ex As System.Exception
                Return -1
            End Try
        End Function


        ''' <summary>
        ''' Convete Point para Milímetro
        ''' </summary>
        ''' <param name="Value">Valor em point</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PtToMM(ByVal Value As Single) As Single
            Try
                Return ((Value * 20) / 56.7)
            Catch ex As System.Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Convete Point para Twips
        ''' </summary>
        ''' <param name="Value">Valor em point</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PtToTw(ByVal Value As Single) As Single
            Try
                Return (Value * 20)
            Catch ex As System.Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Converte Twips para Point
        ''' </summary>
        ''' <param name="Value">Valor em twip</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TwToPt(ByVal Value As Single) As Single
            Try
                Return (Value / 20)
            Catch ex As System.Exception
                Return -1
            End Try
        End Function


        ''' <summary>
        ''' Convete Pixel para Milímetro
        ''' </summary>
        ''' <param name="Value">Valor em pixel</param>
        ''' <param name="DpiPos">Eixo DPI</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PxToMm(ByVal Value As Single, ByVal DpiPos As DpiPos) As Single
            Try
                Dim MyGr As Graphics = Graphics.FromImage(New Bitmap(1, 1, Imaging.PixelFormat.Format32bppRgb))
                Select Case DpiPos
                    Case DpiPos.X
                        Return (((2.54 * Value) / MyGr.DpiX) * 10)
                    Case DpiPos.Y
                        Return (((2.54 * Value) / MyGr.DpiY) * 10)
                    Case Else
                        Return Nothing
                End Select
            Catch ex As System.Exception
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Converte Point em Pixel
        ''' </summary>
        ''' <param name="Value">Valor em point</param>
        ''' <param name="DpiPos">Eixo DPI</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PtToPx(Value As Single, ByVal DpiPos As DpiPos) As Single
            Return MmToPx(PtToMM(Value), DpiPos)
        End Function

    End Class

End Namespace