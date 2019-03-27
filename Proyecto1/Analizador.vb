Imports itextsharp.text.pdf
Imports itextsharp
Imports System.IO
Imports itextsharp.text

Public Class Analizador
    Public Shared Sub AnalisisLexico(v As String, nomArchivo As String)
        Dim errorcito As Boolean = False
        Dim estado As Integer = 0
        Dim lexema As String = ""
        Dim letra As Char
        Dim aski As Integer
        Dim indice As Integer = 0
        Dim comentario As Boolean = False

        Dim cadena As Boolean = False
        Dim Texto As String = ""
        Dim cadenaT As Boolean = False
        Dim TextoTachado As String = ""
        Dim cadenaN As Boolean = False
        Dim TextoNegrita As String = ""
        Dim posibleCadena As Boolean = False

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

        'Inicio del Analizador
        For indice = 0 To v.Length - 1

            letra = v.Chars(indice)
            aski = Asc(letra)

            If aski = 10 Then
                If (errorcito = True) Then
                    fila += 1
                    columna = 1
                End If
            ElseIf (aski = 32 And cadena = False And cadenaT = False And cadenaN = False) Then
                '---Omite espacios en blanco
                columna = columna + 1
            Else
                Select Case estado

                    Case 0
                        errorcito = False
                        'Reconoce los nombres de Variables
                        If (aski >= 65 And aski <= 90) Or (aski >= 97 And aski <= 122) Or aski = 95 Then
                            estado = 1
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 45 Then 'Reconoce el signo menos
                            estado = 2
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski >= 48 And aski <= 57 Then 'Reconoce los´números
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 123 Then 'Reconoce la llave de Inicio {
                            estado = 4
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 125 Then 'Reconoce la llave de Fin }
                            estado = 5
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 40 Then 'Reconoce (
                            estado = 6
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 41 Then 'Reconoce )
                            estado = 7
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 58 Then 'Reconoce :
                            estado = 8
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 34 Then 'Reconoce "
                            estado = 9
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 61 Then 'Reconoce =
                            estado = 11
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 47 Then 'Reconoce /
                            estado = 12
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 42 Then 'Reconoce *
                            estado = 13
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 59 Then 'Reconoce ;
                            estado = 14
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 44 Then 'Reconoce ,
                            estado = 15
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 91 Then 'Reconoce [
                            estado = 16
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 93 Then 'Reconoce ]
                            estado = 17
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 43 Then 'Reconoce +
                            estado = 18
                            lexema = lexema + letra
                            longitudPalabra += 1

                        Else
                            errorcito = True
                            lexema = lexema + letra
                            longitudPalabra += 1
                            errorReconocido.Add(lexema)

                            filaError.Add(fila)
                            columnaError.Add(columna)

                            columna = columna + longitudPalabra
                            lexema = ""
                            estado = 0
                        End If

                    Case 1


                        If (cadena = True) Then
                            'Si lo que voy a evaluar es cadena, permitiré todos los caracteres menos "
                            If (aski >= 32 And aski <= 33) Or (aski >= 35 And aski <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            ElseIf (aski = 34) Then
                                'Aquí ha leido "
                                estado = 9
                                lexema = lexema + letra
                                longitudPalabra += 1
                            End If

                        ElseIf (cadenaT = True) Then
                            'Aquí evaluo la cadena Tachada y no permito el caracter *
                            If (aski >= 32 And aski <= 41) Or (aski >= 43 And aski <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1

                            ElseIf (aski = 42) Then
                                'Aquí ha leido *
                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)
                                tipoToken.Add(token.retornarToken("TEXTO"))

                                '---Escribo el lexema en el PDF con su respectivo Tachado

                                '---
                                cadenaT = False
                                estado = 0
                                columna = columna + longitudPalabra
                                longitudPalabra = 0
                                indice = indice - 1
                                lexema = ""
                                posibleCadena = False
                            End If
                        ElseIf (cadenaN = True) Then
                            If (aski >= 32 And aski <= 42) Or (aski >= 44 And aski <= 254) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            ElseIf (aski = 43) Then
                                '---Fin de cadena al encontrar
                                'el simbolo +
                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)
                                tipoToken.Add(token.retornarToken("TEXTO"))
                                '---

                                '---Escribir Lexema en Negrita en el PDF

                                '---
                                cadenaN = False
                                estado = 0
                                columna = columna + longitudPalabra - 1
                                longitudPalabra = 0
                                indice = indice - 1
                                lexema = ""
                                posibleCadena = False
                            End If
                        Else
                            'Si lo que viene es una letra o un guion bajo _ pero no forma parte de una cadena.
                            'Pueden ser nombres de Variables o Palabras Reservadas
                            If (aski >= 65 And aski <= 90) Or (aski >= 97 And aski <= 122) Or aski = 95 Or (aski >= 48 And aski <= 57) Then
                                estado = 1
                                lexema = lexema + letra
                                longitudPalabra += 1
                            Else

                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)

                                If (String.Compare("INSTRUCCIONES", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                ElseIf (String.Compare("VARIABLES", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                ElseIf (String.Compare("TEXTO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                ElseIf (String.Compare("INTERLINEADO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                ElseIf (String.Compare("TAMANIO_LETRA", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("NOMBRE_ARCHIVO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("DIRECCION_ARCHIVO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("IMAGEN", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("NUMEROS", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                ElseIf (String.Compare("LINEA_EN_BLANCO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("VAR", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("PROMEDIO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("SUMA", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("ASIGNAR", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("CADENA", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))


                                ElseIf (String.Compare("ENTERO", lexema.ToUpper)) = 0 Then
                                    tipoToken.Add(token.retornarToken("PALABRA"))

                                Else
                                    tipoToken.Add(token.retornarToken("ID"))

                                End If

                                estado = 0
                                columna = columna + longitudPalabra
                                longitudPalabra = 0
                                lexema = ""
                                indice = indice - 1
                            End If
                        End If

                    Case 2
                        'Acepta -

                        '-------------------Verifica que lo que venga sean números ya que antes hubo un signo menos - ----------
                        If aski >= 97 And aski <= 122 Then
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            'No vino ningún número, se toma como error
                            '---------------------------------------------
                            errorReconocido.Add(lexema)
                            filaError.Add(fila)
                            columnaError.Add(columna)

                            'tipoToken.Add(token.retornarToken("ERROR"))
                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                            '---------------------------------------------
                        End If
                    Case 3
                        'Verifica que venga un número dado que antes hay un número
                        If aski >= 48 And aski <= 57 Then
                            estado = 3
                            lexema = lexema + letra
                            longitudPalabra += 1

                        ElseIf aski = 46 Then
                            '---------Vino punto, entonces es número Real------------
                            estado = 20
                            lexema = lexema + letra
                            longitudPalabra = longitudPalabra + 1

                        ElseIf (aski >= 65 And aski <= 126) Then
                            'Tengo que enviarlo a un estado de Error

                            estado = 25
                            lexema = lexema + letra
                            longitudPalabra += 1

                        Else
                            'Vino un elemento distinto a número, pero es estado de aceptación se reconoce como lexema.
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("NUMERO"))
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
                        tipoToken.Add(token.retornarToken("{"))

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
                        tipoToken.Add(token.retornarToken("}"))

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
                        tipoToken.Add(token.retornarToken("("))

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
                        tipoToken.Add(token.retornarToken(")"))

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
                        tipoToken.Add(token.retornarToken(":"))

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
                            If (aski >= 32 And aski <= 33) Or (aski >= 35 And aski <= 254) Then
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
                            tipoToken.Add(token.retornarToken("TEXTO"))
                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        Else
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken(Chr(34)))

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
                        tipoToken.Add(token.retornarToken("="))

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 12
                        'Acepta /
                        If (aski = 42) Then
                            estado = 13
                            comentario = True
                        Else
                            '---Acepta diagonal sin asterisco, entonces no es comentario.
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("/"))

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 13
                        'Acepta *
                        If posibleCadena = True Then
                            'Si posibleCadena es verdadera, entonces antes ha venido un [ 
                            'por lo tanto se procederá a guardar TEXTO
                            If (aski >= 33 And aski <= 41) Or (aski >= 43 And aski <= 254) Then
                                estado = 1
                                longitudPalabra += 1
                                cadenaT = True

                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)
                                tipoToken.Add(token.retornarToken("*"))
                                'estado = 0
                                columna = columna + longitudPalabra
                                longitudPalabra = 0
                                lexema = ""
                                indice = indice - 1
                                posibleCadena = False
                            End If
                        ElseIf comentario = True Then
                            If (aski >= 32 And aski <= 41) Or (aski >= 43 And aski <= 254) Then
                                estado = 13
                                columna += 1
                            Else
                                comentario = False
                                indice += 1
                            End If
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
                        tipoToken.Add(token.retornarToken(";"))

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
                        tipoToken.Add(token.retornarToken(","))

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""

                        indice = indice - 1
                    Case 16
                        'Acepta [
                        If (aski = 42) Then
                            'Entra *, por lo tanto el texto encerrado será tachado
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("["))

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            indice -= 1
                            lexema = ""
                            posibleCadena = True
                        ElseIf (aski = 43) Then
                            'Entra +, por lo tanto el texto encerrado será en negrita
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("["))

                            estado = 0
                            columna = columna + longitudPalabra - 1
                            longitudPalabra = 0
                            indice = indice - 1
                            lexema = ""
                            posibleCadena = True
                        Else
                            posibleCadena = False
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("["))

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 17
                        'Acepta ]
                        lexemaReconocido.Add(lexema)
                        filaLexema.Add(fila)
                        columnaLexema.Add(columna)
                        tipoToken.Add(token.retornarToken("]"))

                        estado = 0
                        columna = columna + longitudPalabra
                        longitudPalabra = 0
                        lexema = ""
                        indice = indice - 1
                    Case 18
                        'Acepta +

                        '---Verifica si posibleCadena está habilitado---
                        '---Si está habilitado, entonces antes vino un [
                        If (posibleCadena = True) Then
                            If (aski >= 33 And aski <= 42) Or (aski >= 44 And aski <= 254) Then
                                estado = 1
                                longitudPalabra += 1
                                cadenaN = True

                                lexemaReconocido.Add(lexema)
                                filaLexema.Add(fila)
                                columnaLexema.Add(columna)
                                tipoToken.Add(token.retornarToken("+"))
                                'estado = 0
                                columna = columna + longitudPalabra
                                lexema = ""
                                indice = indice - 1
                                posibleCadena = False
                            End If
                        Else
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)

                            tipoToken.Add(token.retornarToken("+"))


                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If

                    Case 20
                        'Verifica que venga número ya que ha venido punto
                        If aski >= 48 And aski <= 57 Then
                            estado = 21
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
                        If aski >= 48 And aski <= 57 Then
                            estado = 21
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            'No viene otro número, pero lo acepta ya que es estado de aceptación.
                            lexemaReconocido.Add(lexema)
                            filaLexema.Add(fila)
                            columnaLexema.Add(columna)
                            tipoToken.Add(token.retornarToken("RNUMERO"))

                            estado = 0
                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                    Case 25
                        If (aski >= 48 And aski <= 57) Or (aski >= 65 And aski <= 90) Or (aski >= 97 And aski <= 122) Then
                            estado = 25
                            lexema = lexema + letra
                            longitudPalabra += 1
                        Else
                            estado = 0
                            errorReconocido.Add(lexema)
                            filaError.Add(fila)
                            columnaError.Add(columna)

                            columna = columna + longitudPalabra
                            longitudPalabra = 0
                            lexema = ""
                            indice = indice - 1
                        End If
                End Select
                If indice > 0 Then
                    If (Asc(v.Chars(indice - 1)) = 10) Then
                        fila += 1
                        columna = 1
                        If (errorcito = True) Then
                            fila -= 1
                            If (columna > 1) Then
                                columna = 1
                            Else
                                columna += 1
                            End If
                        End If
                    End If
                End If
            End If

        Next


        '---Creación PDF
        If (errorReconocido.Count > 0) Then
            MessageBox.Show("El archivo contiene errores." & vbLf& & vbLf & "En su escritorio hallará un archivo, donde se especifica que errores hay en el Archivo.", "Analizador Léxico v0.1")
            PDF.crearPDF_Errores(errorReconocido, nomArchivo, filaError, columnaError)
        Else
            MessageBox.Show("El archivo ha sido analizado correctamente. " & vbLf & "Observará un documento en su Escritorio" & vbLf & "Sobre el Análisis Léxico Realizado", "Analizador Léxico v0.1")
            '---Reconocimiento de Lexemas para realizar el apartado de instrucciones sin la evaluación correspondiente.
            Dim archivoSalida As String = "Default.pdf"
            Dim interlineado As Double = 1.5
            Dim dirArchivo As String = "C:\"
            Dim sizeLetra As Integer = 11

            For i = 0 To lexemaReconocido.Count - 1
                If (lexemaReconocido(i).ToString.ToUpper = "nombre_archivo".ToUpper) Then
                    '---Nombre de Archivo si exite...

                    archivoSalida = lexemaReconocido(i + 2).ToString
                    archivoSalida = archivoSalida.Replace(Chr(34), "")
                    'i = lexemaReconocido.Count + 1
                End If
                If (lexemaReconocido(i).ToString.ToUpper = "Interlineado".ToUpper) Then
                    interlineado = Convert.ToDouble(lexemaReconocido(i + 2).ToString)

                End If
                If (lexemaReconocido(i).ToString.ToUpper = "Tamanio_letra".ToUpper) Then
                    sizeLetra = Convert.ToInt32(lexemaReconocido(i + 2).ToString)
                End If
                If (lexemaReconocido(i).ToString.ToUpper = "direccion_archivo".ToUpper) Then
                    dirArchivo = lexemaReconocido(i + 2).ToString
                End If
            Next
            '---Paso de Parámetros para crear el PDF de Usuario.
            PDF.crearPDFUser(lexemaReconocido, interlineado, sizeLetra, dirArchivo, archivoSalida)
            PDF.crearPDF(lexemaReconocido, tipoToken, nomArchivo, archivoSalida, filaLexema, columnaLexema)
        End If

    End Sub
End Class