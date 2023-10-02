'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.IO
Imports System.Drawing
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Drawing.Imaging
Imports acApp = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports System.Runtime.InteropServices
Imports System.Windows.Interop
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Colors
Imports System.Drawing.Printing
Imports System.Drawing.Drawing2D
Imports System.Collections
Imports System.Configuration

Namespace Engine2

    ''' <summary>
    ''' Permite recuperar a imagem preview do DWG
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class FileThumbnail

        ''' <summary>
        ''' Eixo de espelhamento
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eReverseAxisImage
            ''' <summary>
            ''' Espelhamento em X
            ''' </summary>
            ''' <remarks></remarks>
            X = 0

            ''' <summary>
            ''' Espelhamento em Y
            ''' </summary>
            ''' <remarks></remarks>
            Y = 1
        End Enum

        ''' <summary>
        ''' Retorna a imagem do DWG (Utiliza API do AutoCAD)
        ''' </summary>
        ''' <param name="Path">Caminho do arquivo de desenho</param>
        ''' <param name="TransparentBackColor">Determina se o fundo da imagem será transaparente</param>
        ''' <returns>Bitmap</returns>
        ''' <remarks></remarks>
        Public Shared Function GetThumbnail(Path As Object, Optional TransparentBackColor As Boolean = False) As Bitmap
            Dim Bitmap As Bitmap = Nothing
            Dim color As System.Drawing.Color
            If IsNothing(Path) = False Then
                Try
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using DocumentLock As DocumentLock = Document.LockDocument()
                        Using Database As New Database(False, False)
                            With Database
                                .ReadDwgFile(Path, FileOpenMode.OpenForReadAndReadShare, False, "")
                                Bitmap = .ThumbnailBitmap
                                .Dispose()
                            End With
                        End Using
                    End Using
                    If TransparentBackColor = True Then
                        color = Bitmap.GetPixel(0, 0)
                        Bitmap.MakeTransparent(color)
                    End If
                    Return Bitmap
                Catch ex As System.Exception
                    Return Bitmap
                End Try
            Else
                Return Bitmap
            End If
        End Function

        ''' <summary>
        ''' Obtem a imagem do arquivo DWG (Não utiliza API do AutoCAD)
        ''' </summary>
        ''' <param name="Path">Caminho do arquivo de desenho</param>
        ''' <param name="TransparentBackColor">Determina se o fundo da imagem será transaparente</param>
        ''' <returns>Bitmap</returns>
        ''' <remarks></remarks>
        Public Shared Function GetThumbnail2(Path As String, Optional TransparentBackColor As Boolean = False) As Bitmap

            'Seta o retorno
            GetThumbnail2 = Nothing

            Try

                'Declarações
                Dim BinaryReader As BinaryReader
                Dim imgPos As Int32
                Dim imgBSentinel() As Byte
                Dim imgCSentinel(16) As Byte
                Dim imgSize As UInt32
                Dim imgPresent As Byte
                Dim imgHeaderStart As Int32
                Dim imgHeaderSize As Int32
                Dim imgBmpStart As Int32
                Dim imgBmpSize As Int32
                Dim bmpDataPresent As Boolean
                Dim imgWmfStart As Int32
                Dim imgWmfSize As Int32
                Dim wmfDataPresent As Boolean
                Dim tempPixelData() As Byte
                Dim tempBuffData() As Byte
                Dim MemoryStream As MemoryStream
                Dim BitMap As Bitmap
                Dim imgCode As Byte
                Dim color As System.Drawing.Color

                Using FileStream = New FileStream(Path, FileMode.Open, IO.FileAccess.Read, FileShare.ReadWrite)
                    BinaryReader = New BinaryReader(FileStream)
                    BinaryReader.BaseStream.Seek(&HD, SeekOrigin.Begin)
                    imgPos = BinaryReader.ReadInt32()
                    BinaryReader.BaseStream.Seek(imgPos, SeekOrigin.Begin)
                    imgBSentinel = {&H1F, &H25, &H6D, &H7, &HD4, &H36, &H28, &H28, &H9D, &H57, &HCA, &H3F, &H9D, &H44, &H10, &H2B}
                    imgCSentinel = BinaryReader.ReadBytes(16)
                    If (imgBSentinel.ToString() = imgCSentinel.ToString()) Then
                        imgSize = BinaryReader.ReadUInt32()
                        imgPresent = BinaryReader.ReadByte()
                        imgHeaderStart = 0
                        imgHeaderSize = 0
                        imgBmpStart = 0
                        imgBmpSize = 0
                        bmpDataPresent = False
                        imgBmpStart = 0
                        imgBmpSize = 0
                        bmpDataPresent = False
                        wmfDataPresent = False
                        For I As Integer = 1 To imgPresent
                            imgCode = BinaryReader.ReadByte()
                            Select Case imgCode
                                Case 1
                                    'DATA
                                    imgHeaderStart = BinaryReader.ReadInt32()
                                    imgHeaderSize = BinaryReader.ReadInt32()
                                Case 2
                                    'BMP
                                    imgBmpStart = BinaryReader.ReadInt32()
                                    imgBmpSize = BinaryReader.ReadInt32()
                                    bmpDataPresent = True
                                Case 3
                                    'WPF
                                    imgWmfStart = BinaryReader.ReadInt32()
                                    imgWmfSize = BinaryReader.ReadInt32()
                                    wmfDataPresent = True
                            End Select
                        Next
                    End If
                    If (bmpDataPresent) Then
                        BinaryReader.BaseStream.Seek(imgBmpStart, SeekOrigin.Begin)
                        ReDim tempPixelData(imgBmpSize + 14)
                        tempPixelData(0) = &H42
                        tempPixelData(1) = &H4D
                        tempPixelData(10) = &H36
                        tempPixelData(11) = &H4
                        ReDim tempBuffData(imgBmpSize)
                        tempBuffData = BinaryReader.ReadBytes(imgBmpSize)
                        tempBuffData.CopyTo(tempPixelData, 14)
                        MemoryStream = New MemoryStream(tempPixelData)
                        BitMap = New Bitmap(MemoryStream)
                        If TransparentBackColor = True Then
                            color = BitMap.GetPixel(0, 0)
                            BitMap.MakeTransparent(color)
                        End If
                        GetThumbnail2 = BitMap
                    End If
                    FileStream.Close()
                End Using

            Catch

                'Mata o erro
                Exit Try

            End Try

            'Return
            Return GetThumbnail2

        End Function

        ''' <summary>
        ''' Rotaciona um bitmap
        ''' </summary>
        ''' <param name="Image">Image</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RotateBitmap(Image As Bitmap, DegreeAngle As Single) As Bitmap
            Dim X As Single
            Dim Y As Single
            Dim BitmapClone As New Bitmap(Image)
            Dim Width As Single = BitmapClone.Width
            Dim Height As Single = BitmapClone.Height
            Dim Corners As Point() = {New Point(0, 0), New Point(Width, 0), New Point(0, Height), New Point(Width, Height)}
            Dim CornerX As Single = (Width / 2)
            Dim CornerY As Single = (Height / 2)
            For Index As Integer = 0 To 3
                Corners(Index).X -= CornerX
                Corners(Index).Y -= CornerY
            Next
            Dim Theta As Single = Single.Parse(DegreeAngle) * Math.PI / 180.0
            Dim SinTheta As Single = Math.Sin(Theta)
            Dim CosTheta As Single = Math.Cos(Theta)
            For Index As Integer = 0 To 3
                X = Corners(Index).X
                Y = Corners(Index).Y
                Corners(Index).X = X * CosTheta + Y * SinTheta
                Corners(Index).Y = -X * SinTheta + Y * CosTheta
            Next
            Dim Xmin As Single = Corners(0).X
            Dim Ymin As Single = Corners(0).Y
            For Index As Integer = 1 To 3
                If Xmin > Corners(Index).X Then Xmin = Corners(Index).X
                If Ymin > Corners(Index).Y Then Ymin = Corners(Index).Y
            Next
            For Index As Integer = 0 To 3
                Corners(Index).X -= Xmin
                Corners(Index).Y -= Ymin
            Next
            Dim BitmapOut As New Bitmap(CInt(-2 * Xmin), CInt(-2 * Ymin))
            Dim GraphicsOut As Graphics = Graphics.FromImage(BitmapOut)
            ReDim Preserve Corners(2)
            GraphicsOut.DrawImage(BitmapClone, Corners)
            Return BitmapOut
        End Function

        ''' <summary>
        ''' Reverte a imagem
        ''' </summary>
        ''' <param name="Image"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReverseImage(Image As Bitmap, ReverseAxis As eReverseAxisImage) As Bitmap
            Select Case ReverseAxis
                Case eReverseAxisImage.X
                    Image.RotateFlip(RotateFlipType.RotateNoneFlipX)
                Case eReverseAxisImage.Y
                    Image.RotateFlip(RotateFlipType.RotateNoneFlipY)
            End Select
            Return Image
        End Function

    End Class

End Namespace



