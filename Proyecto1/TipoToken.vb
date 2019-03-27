Public Class TipoToken
    Friend Function retornarToken(v As String) As Object
        Select Case v
            Case "PALABRA"
                Return "Palabra Reservada"
            Case "ID"
                Return "Identificador"
            Case "TEXTO"
                Return "Cadena"
            Case "NUMERO"
                Return "Número"
            Case "RNUMERO"
                Return "Número Real"
            Case "{"
                Return "Llave de Inicio"
            Case "}"
                Return "Llave de Cierre"
            Case "("
                Return "Paréntesis de Inicio"
            Case ")"
                Return "Paréntesis de Cierre"
            Case ":"
                Return "Dos Puntos"
            Case Chr(34)
                Return "Comilla"
            Case "/"
                Return "Barra"
            Case "*"
                Return "Asterisco"
            Case ";"
                Return "Punto y Coma"
            Case "["
                Return "Corchete de Inicio"
            Case "]"
                Return "Corchete de Cierre"
            Case "+"
                Return "Cruz"
            Case ","
                Return "Coma"
            Case Else
                Return "Palabra no identificada"
        End Select
    End Function
End Class
