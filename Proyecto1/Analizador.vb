Imports itextsharp.text.pdf
Imports itextsharp.text
Imports System.IO
Public Class Analizador
    Public Shared Sub AnalisisLexico(v As String)
        Dim estado As Integer = 0
        Dim lexema As String = ""
        Dim letra As Char
        Dim codigoAscii As Integer
        Dim indice As Integer = 0

        Dim cadena As Boolean = False
        Dim cadenaT As Boolean = False
        Dim cadenaN As Boolean = False

        Dim lexemaReconocido As New ArrayList
        Dim filaLexema As New ArrayList
        Dim columnaLexema As New ArrayList

        Dim errorReconocido As New ArrayList
        Dim filaError As New ArrayList
        Dim columnaError As New ArrayList

        Dim tipoToken As New ArrayList
        Dim token As New TipoToken

        Dim fila, columna, longitudPalabra As Integer
        fila = 1
        columna = 1
        longitudPalabra = 0

        For indice = 1 To v.Length
            letra = v.Chars(indice - 1)
            codigoAscii = Asc(letra)
            If (codigoAscii = 10) Then 'Verifica Saltos de Línea
                fila = fila + 1
                columna = 0
            ElseIf (codigoAscii = 32 And cadena = False And cadenaT = False) Then 'Verifica Espacios en Blanco solamente cuando no sea cadena
                columna = columna + 1
            Else
                Select Case estado
                    Case 0
                        'Reconoce los nombres de Variables
                        If (codigoAscii >= 65 And codigoAscii <= 90) Or (codigoAscii >= 97 And codigoAscii <= 122) Or codigoAscii = 95 Then
                            estado = 1
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 45 Then 'Reconoce el signo menos
                            estado = 2
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii >= 48 And codigoAscii <= 57 Then 'Reconoce los´números
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 123 Then 'Reconoce la llave de Inicio {
                            estado = 4
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 125 Then 'Reconoce la llave de Fin }
                            estado = 5
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 40 Then 'Reconoce (
                            estado = 6
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 41 Then 'Reconoce )
                            estado = 7
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 58 Then 'Reconoce :
                            estado = 8
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 34 Then 'Reconoce "
                            estado = 9
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 61 Then 'Reconoce =
                            estado = 11
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 47 Then 'Reconoce /
                            estado = 12
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 42 Then 'Reconoce *
                            estado = 13
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 59 Then 'Reconoce ;
                            estado = 14
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 44 Then 'Reconoce ,
                            estado = 15
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 91 Then 'Reconoce [
                            estado = 16
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 93 Then 'Reconoce ]
                            estado = 17
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 43 Then 'Reconoce +
                            estado = 18
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            lexema = lexema + letra
                            errorReconocido.Add(lexema)
                            token.retornarToken("ERROR")
                            lexema = ""
                        End If
                    Case 1
                        'Si lo que voy a evaluar es cadena, permitiré todos los caracteres menos "
                        If (cadena = True) Then
                            If (codigoAscii >= 32 And codigoAscii <= 33) Or (codigoAscii >= 35 And codigoAscii <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            ElseIf (codigoAscii = 34) Then
                                estado = 9
                                lexema = lexema + letra
                                longitudPalabra += 1
                            End If
                        ElseIf (cadenaT = True) Then
                            If (codigoAscii >= 32 And codigoAscii <= 41) Or (codigoAscii >= 43 And codigoAscii <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            ElseIf (codigoAscii = 42) Then '*
                                estado = 13
                                lexema = lexema + letra
                                longitudPalabra += 1
                            End If
                        Else
                            'Si lo que viene es una letra o un guion bajo _ pero no forma parte de una cadena.
                            If (codigoAscii >= 65 And codigoAscii <= 90) Or (codigoAscii >= 97 And codigoAscii <= 122) Or codigoAscii = 95 Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            ElseIf codigoAscii = 46 Then '.
                                estado = 20
                                lexema = lexema + letra
                                longitudPalabra += 1
                            Else
                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)

                                estado = 0

                                columna = columna + longitudPalabra
                                longitudPalabra = 0
                                lexema = ""
                                indice = indice - 1
                            End If
                        End If
                    Case 2
                        'Acepta -
                        'Verifica que lo que venga sean números ya que antes hubo un signo menos - 
                        If codigoAscii >= 97 And codigoAscii <= 122 Then
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            'No vino ningún número, se toma como error
                            errorReconocido.Add(lexema)
                            filaError.Add(fila)
                            columnaError.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 3
                        'Verifica que venga un número dado que antes hay un número
                        If codigoAscii >= 97 And codigoAscii <= 122 Then
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1
                        ElseIf codigoAscii = 46 Then
                            'Verifica si lo que viene es un punto
                            estado = 20
                            lexema = lexema + letra
                            longitudPalabra = longitudPalabra + 1
                        Else
                            'Vino un elemento distinto a número, pero es estado de aceptación se reconoce como lexema.
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 4
                        'Acepta {
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 5
                        'Acepta }
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 6
                        'Acepta (
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 7
                        'Acepta )
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 8
                        'Acepta :
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 9
                        'Acepta "
                        'Reconociendo el Inicio de Cadena
                        If (cadena = False) Then
                            cadena = True
                            If (codigoAscii >= 32 And codigoAscii <= 33) Or (codigoAscii >= 35 And codigoAscii <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            End If
                        ElseIf (cadena = True) Then
                            'Reconociendo el Final de Cadena
                            cadena = False
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 11
                        'Acepta =
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 12
                        'Acepta /
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 13
                        'Acepta *
                        If (cadenaT = False) Then
                            If (codigoAscii >= 32 And codigoAscii <= 41) Or (codigoAscii >= 43 And codigoAscii <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                                cadenaT = True
                            End If
                        ElseIf (cadenaT = True) Then
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("CADENA"))
                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                            cadenaT = False
                        Else
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("*"))
                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If

                    Case 14
                        'Acepta ;
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 15
                        'Acepta ,
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""

                        indice = indice - 1
                    Case 16
                        'Acepta [
                        If (codigoAscii = 42) Then
                            'Entra *, por lo tanto el texto encerrado será tachado
                            estado = 13
                        ElseIf (codigoAscii = 43) Then
                            'Entra +, por lo tanto el texto encerrado será en negrita
                            estado = 18
                        End If
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 17
                        'Acepta ]
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 18
                        'Acepta +
                        If (codigoAscii >= 32 And codigoAscii <= 42) Or (codigoAscii >= 44 And codigoAscii <= 254) Then
                            estado = 1
                            lexema = lexema + letra
                            longitudPalabra += 1
                            cadena = True
                        Else
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If

                    Case 20
                        'Verifica que venga número ya que ha venido punto
                        If codigoAscii >= 97 And codigoAscii <= 122 Then
                            estado = 21
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf codigoAscii >= 65 And codigoAscii <= 90 Or (codigoAscii >= 97 And codigoAscii <= 122) Or codigoAscii = 95 Then
                            'Verifica que venga alguna letra ya que ha venido punto, para representar una extensión...
                            estado = 1
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            'No viene ni letra ni número, por lo tanto el lexema capturado es considerado error.
                            errorReconocido.Add(lexema)
                            filaError.Add(fila)
                            columnaError.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 21
                        'Aceptador de numeros reales
                        If codigoAscii >= 97 And codigoAscii <= 122 Then
                            estado = 21
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            'No viene otro número, pero lo acepta ya que es estado de aceptación.
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                        End If
                End Select
            End If
        Next

        recorrerLexemas(lexemaReconocido)

        recorrerErrores(errorReconocido)

    End Sub

    Private Shared Sub recorrerErrores(errorReconocido As ArrayList)
        Dim rechazado As String
        For Each rechazado In errorReconocido
            MessageBox.Show(rechazado, "Errores")
        Next
    End Sub

    Private Shared Sub recorrerLexemas(lexemaReconocido As ArrayList)
        Dim lexema As String
        For Each lexema In lexemaReconocido
            MessageBox.Show(lexema, "Lexemas")
        Next
    End Sub
End Class
