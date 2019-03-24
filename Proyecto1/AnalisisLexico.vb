Public Class AnalisisLexico
    Friend Shared Sub AnalisisLexico(v As String)
        'Variable apuntador
        Dim indice As Integer = 0
        Dim letra As String = ""
        'Codigo ASCII del caracter
        Dim aski As Integer = 0


        'Variable para reconocer Lexemas
        Dim lexema As String = ""
        'Variable para reconocer en que estado estoy
        Dim estado As Integer = 0


        'ArrayList para guardar lexemas
        Dim listaLexemas As New ArrayList
        'ArrayList para guardar errores
        Dim listaErrores As New ArrayList

        For indice = 0 To v.Length - 1
            aski = Asc(v.Chars(indice))

            If (aski = 10) Then
                'Verifica Saltos de Línea
            ElseIf (aski = 32) Then
                'Verifica Espacios en Blanco
            Else
                'Sino viene espacio en blanco o salto de línea, viene una letra. Hay que evaluarla


                'Vamos a obtener los tokens tipo identificador y número
                Select Case estado
                'De el estado 0 parten todos los demás estados. No es aceptador
                    Case 0
                        'Definiré el intervalo aski para aceptar letras también podrán contener guion bajo
                        If (aski >= 65 And aski <= 90) Or (aski >= 97 And aski <= 122) Or (aski = 95) Then
                            estado = 1
                            lexema = lexema + v.Chars(indice)
                        End If
                    Case 1
                        'Con esta condicional aceptaré numeros, letras minusculas mayusculas y guion bajo en el nombre más no letras acentuadas
                        If (aski >= 65 And aski <= 90) Or (aski >= 97 And aski <= 122) Or (aski = 95) Or (aski >= 48 And aski <= 57) Then
                            estado = 1
                            'Si eso sucede, tengo que guardar letra por letra
                            lexema = lexema + v.Chars(indice)
                        Else
                            'Cuando eso no ocurra, eso quiere decir que ha venido un caracter fuera de esos límites.
                            'Este estado de aceptación, por lo tanto tendré que registrar este token.
                            listaLexemas.Add(lexema)
                            'Vacio el lexema para dar lugar a otro
                            lexema = ""
                        End If
                    Case Else
                        'Si no cumple con ninguno de los estados, entonces es un error 
                        listaErrores.Add(lexema)
                        lexema = ""
                End Select
            End If
        Next
    End Sub
End Class
