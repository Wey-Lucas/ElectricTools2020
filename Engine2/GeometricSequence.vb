Imports System.Text
Imports System.Numerics

Namespace Engine2

    ''' <summary>
    ''' Gerência sequencias geométricas. 
    ''' Permite criar seleções que resultam em um número unico. 
    ''' Do número gerado é possível retornar a lista de de itens que originalmente foi selecionado. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class GeometricSequence

        ''' <summary>
        ''' Instancia Sequences. 
        ''' </summary>
        ''' <remarks></remarks>
        Private _Sequences As Sequences

        ''' <summary>
        ''' Permite a inclusão ou consulta de sequencias. 
        ''' </summary>
        ''' <value></value>
        ''' <returns>Sequences</returns>
        ''' <remarks></remarks>
        Public Property Sequences() As Sequences
            Get
                Return _Sequences
            End Get
            Set(ByVal value As Sequences)
                _Sequences = value
            End Set
        End Property

        ''' <summary>
        ''' Retorna os itens dado a sequencia numérica.
        ''' </summary>
        ''' <param name="NumericSequence">Sequencia numérica.</param>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Public Function GetItemsInSequence(ByVal NumericSequence As BigInteger) As ArrayList
            Try

                'Declara as variáveis
                Dim arrLst As New ArrayList
                Dim Text As String
                Dim Value As BigInteger
                Dim NextValue As BigInteger = 0
                Dim dRem As BigInteger = 0

                'Monta  a lista de saída
                With _Sequences
                    While NumericSequence <> 0
                        For ini As BigInteger = (_Sequences.Count - 1) To 0 Step -1
                            Text = .Item(ini).Text
                            Value = .Item(ini).Value
                            If Value <= NumericSequence Then
                                arrLst.Add(Text)
                                NumericSequence = NumericSequence - Value
                            End If
                            If NumericSequence = 0 Then
                                Exit For
                            End If
                        Next
                    End While
                End With

                'Ordena o resultado
                arrLst.Sort()

                'Retorna o valor
                Return arrLst

            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Retorna a soma de itens.
        ''' </summary>
        ''' <param name="lstTextItens">ArrayList com os itens a serem somados.</param>
        ''' <returns>BigInteger</returns>
        ''' <remarks></remarks>
        Public Function SumItemsInSequence(ByVal lstTextItens As ArrayList) As BigInteger
            Try

                'Declarações
                Dim Soma As BigInteger = 0

                'Monta a soma 
                With _Sequences
                    For ini As BigInteger = 0 To (lstTextItens.Count - 1)
                        For ini2 As BigInteger = 0 To (.Count - 1)
                            If .Item(ini2).Text.ToUpper = lstTextItens.Item(ini).ToString.ToUpper Then
                                Soma = Soma + .Item(ini2).Value
                                Exit For
                            End If
                        Next
                    Next
                End With

                'Retorna o valor
                Return Soma

            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Retorna o próximo valor da sequencia.
        ''' </summary>
        ''' <returns>BigInteger</returns>
        ''' <remarks></remarks>
        Public Function GetNextValue() As BigInteger
            Try
                If _Sequences.Count = 0 Then
                    Return 1
                Else
                    Return (_Sequences.Item(_Sequences.Count - 1).Value * 2)
                End If
            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Retorna o código para a criação do enumerador.
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Function CodeToEnum(ByVal EnumName As String) As String
            Try

                'Declarações
                Dim SB As New StringBuilder

                'Validação
                If _Sequences.Count = 0 Then
                    Throw New System.Exception("Não é possível gerar o código para o enumerador." & vbCrLf & "Nenhuma sequencia foi informada.")
                End If

                'Monta o código para o enumerador
                With SB
                    .AppendLine("Public Enum " & EnumName)
                    For ini As BigInteger = 0 To (_Sequences.Count - 1)
                        .AppendLine("   " & _Sequences.Item(ini).Text & " = " & _Sequences.Item(ini).Value.ToString)
                    Next
                    .AppendLine("End Enum")
                End With

                'Retorno
                Return SB.ToString

            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me._Sequences = New Sequences
        End Sub

    End Class

    ''' <summary>
    ''' Classe principal para campos.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Sequences

        'Declarações Inherits
        Inherits List(Of Sequence)

        ''' <summary>
        ''' Adiciona uma sequência numérica associada a um valor textual
        ''' </summary>
        ''' <param name="Text">Texto do enumerador</param>
        ''' <param name="Value">Valor do enumerador</param>
        ''' <returns>Sequence</returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal Text As String, ByVal Value As BigInteger) As Sequence
            Try
                If MyBase.Count > 0 Then
                    If Value <= MyBase.Item(MyBase.Count - 1).Value Then
                        Throw New System.Exception("Não é possível incluir a sequência 'Text = " & Text & " \ Value = " & Value.ToString & "'." & vbCrLf & "O valor é igual ou menor que o valor da última sequência informada.")
                    End If
                    If Value <> (MyBase.Item(MyBase.Count - 1).Value * 2) Then
                        Throw New System.Exception("Não é possível incluir a sequência 'Text = " & Text & " \ Value = " & Value.ToString & "'." & vbCrLf & "O valor não representa o dobro do último valor informado." & vbCrLf & "Valor informado = " & MyBase.Item(MyBase.Count - 1).Value.ToString & "." & vbCrLf & "Valor esperado = " & (MyBase.Item(MyBase.Count - 1).Value.ToString * 2) & ".")
                    End If
                    For ini As BigInteger = 0 To (MyBase.Count - 1)
                        If Text.ToUpper = MyBase.Item(ini).Text.ToUpper Then
                            Throw New System.Exception("Não é possível incluir a sequência 'Text = " & Text & " \ Value = " & Value.ToString & "'." & vbCrLf & "O texto informado já existe.")
                        End If
                    Next
                Else
                    If Value <> 1 Then
                        Throw New System.Exception("Não é possível incluir a sequência 'Text = " & Text & " \ Value = " & Value.ToString & "'." & vbCrLf & "O primeiro valor informado deve ser 1.")
                    End If
                End If
                Dim SequenceItem As New Sequence(Text, Value)
                MyBase.Add(SequenceItem)
                Return SequenceItem
            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Adiciona uma sequência numérica baseando-se em uma arraylist.
        ''' Os valores serão criados automaticamente a partir do número 1.
        ''' </summary>
        ''' <param name="lstText"></param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal lstText As ArrayList) As Boolean
            Try
                Dim Value As BigInteger = 0
                For ini As BigInteger = 0 To (lstText.Count - 1)
                    If MyBase.Count = 0 Then
                        Value = 1
                    Else
                        Value = (MyBase.Item(MyBase.Count - 1).Value * 2)
                    End If
                    Dim SequenceItem As New Sequence(lstText.Item(ini).ToString, Value)
                    If MyBase.Contains(SequenceItem) = True Then
                        Throw New System.Exception("Não é possível incluir a sequência 'Text = " & lstText.Item(ini).ToString & " \ Value = " & Value.ToString & "'." & vbCrLf & "O texto informado já existe.")
                    End If
                    MyBase.Add(SequenceItem)
                Next
                Return True
            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Adiciona uma sequência numérica baseando-se em um enumerador.
        ''' Os valores serão criados automaticamente a partir do número 1.
        ''' Exemplo para passagem de parâmetro: Add(GetType(Nome do enum)).
        ''' </summary>
        ''' <param name="EnumType"></param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal EnumType As System.Type) As Boolean
            Try
                Dim lstText() As String = System.Enum.GetNames(EnumType)
                Dim Value As BigInteger = 0
                For ini As BigInteger = 0 To (lstText.GetLength(0) - 1)
                    If MyBase.Count = 0 Then
                        Value = 1
                    Else
                        Value = (MyBase.Item(MyBase.Count - 1).Value * 2)
                    End If
                    Dim SequenceItem As New Sequence(lstText(ini).ToString, Value)
                    If MyBase.Contains(SequenceItem) = True Then
                        Throw New System.Exception("Não é possível incluir a sequência 'Text = " & lstText(ini).ToString & " \ Value = " & Value.ToString & "'." & vbCrLf & "O texto informado já existe.")
                    End If
                    MyBase.Add(SequenceItem)
                Next
                Return True
            Catch ex As System.Exception
                Throw ex
            End Try
        End Function

    End Class

    ''' <summary>
    ''' Classe para campo.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Sequence

        ''' <summary>
        ''' Armazena o nome textual.
        ''' </summary>
        ''' <remarks></remarks>
        Private oText As String

        ''' <summary>
        ''' Armazena o valor numérico.
        ''' </summary>
        ''' <remarks></remarks>
        Private oValue As BigInteger

        ''' <summary>
        ''' Obtem o texto.
        ''' </summary>
        ''' <value></value>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Text() As String
            Get
                Return oText
            End Get
        End Property

        ''' <summary>
        ''' Obtem o valor do campo.
        ''' </summary>
        ''' <value></value>
        ''' <returns>BigInteger</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value() As BigInteger
            Get
                Return oValue
            End Get
        End Property

        ''' <summary>
        ''' Construtor da sequência.
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Sub New(ByVal Text As String, ByVal Value As BigInteger)
            Try
                oText = Text
                oValue = Value
            Catch ex As System.Exception
                Throw ex
            End Try
        End Sub

    End Class

End Namespace







