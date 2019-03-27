Imports itextsharp.text
Imports itextsharp.text.pdf
Imports System.IO
Public Class PDF
    Friend Shared Sub crearPDF(Lexemas As ArrayList, tipodeToken As ArrayList, nomArchivo As String, archivoSalida As String, filaLexema As ArrayList, columnaLexema As ArrayList)
        Dim pdf As New Document(PageSize.A4)
        Dim escritor As PdfWriter

        escritor = PdfWriter.GetInstance(pdf, New FileStream(My.Computer.FileSystem.SpecialDirectories.Desktop & "/Lexemas.pdf", FileMode.Create, FileAccess.Write, FileShare.None))

        pdf.Open()
        '---Texto Inicial
        pdf.Add(New Paragraph("Universidad de San Carlos de Guatemala"))
        pdf.Add(New Paragraph("Facultad de Ingeniería"))
        pdf.Add(New Paragraph("Escuela de Ciencias"))
        pdf.Add(New Paragraph("Ingeniería en Ciencias y Sistemas"))
        pdf.Add(New Paragraph("Lenguajes Formales y de Programación"))

        '---Instrucciones para agregar Imagen
        Dim logoUsac As Image = Image.GetInstance("Usac_logo.png")
        logoUsac.ScaleToFit(125.0F, 125.0F)
        logoUsac.SetAbsolutePosition(425, 675)
        pdf.Add(logoUsac)
        '---

        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(New Paragraph("Archivo Fuente: " & nomArchivo))
        pdf.Add(New Paragraph("Archivo Salida: " & archivoSalida))
        '---

        '---Tabla con los Lexemas recopilados en el analizador...
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)

        Dim anchoCeldas() As Integer = {1, 10, 6, 1, 2}
        Dim tabla As New PdfPTable(5)
        tabla.WidthPercentage = 100
        tabla.SetWidths(anchoCeldas)
        tabla.AddCell("No.")
        tabla.AddCell("Lexema")
        tabla.AddCell("Tipo")
        tabla.AddCell("Fila")
        tabla.AddCell("Columna")

        '---Recorrido de los ArrayList
        For i = 0 To Lexemas.Count - 1
            tabla.AddCell(i + 1)
            tabla.AddCell(Lexemas.Item(i))
            tabla.AddCell(tipodeToken(i))
            tabla.AddCell(filaLexema(i))
            tabla.AddCell(columnaLexema(i))
        Next
        '---
        pdf.Add(tabla)
        '---
        pdf.Close()
    End Sub

    Friend Shared Sub crearPDF_Errores(errorReconocido As ArrayList, nomArchivo As String, filaerror As ArrayList, columnaError As ArrayList)
        Dim pdf As New Document(PageSize.A4)
        Dim escritor As PdfWriter

        escritor = PdfWriter.GetInstance(pdf, New FileStream(My.Computer.FileSystem.SpecialDirectories.Desktop & "/Errores.pdf", FileMode.Create, FileAccess.Write, FileShare.None))

        pdf.Open()
        '---Texto Inicial
        pdf.Add(New Paragraph("Universidad de San Carlos de Guatemala"))
        pdf.Add(New Paragraph("Facultad de Ingeniería"))
        pdf.Add(New Paragraph("Escuela de Ciencias"))
        pdf.Add(New Paragraph("Ingeniería en Ciencias y Sistemas"))
        pdf.Add(New Paragraph("Lenguajes Formales y de Programación"))

        '---Instrucciones para agregar Imagen
        Dim logoUsac As Image = Image.GetInstance("Usac_logo.png")
        logoUsac.ScaleToFit(125.0F, 125.0F)
        logoUsac.SetAbsolutePosition(425, 675)
        pdf.Add(logoUsac)
        '---

        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(New Paragraph("Archivo Fuente: " & nomArchivo))
        '---

        '---Tabla con los Lexemas recopilados en el analizador...
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)

        Dim anchoCeldas() As Integer = {1, 2, 3, 8, 1, 3}
        Dim tabla As New PdfPTable(6)
        tabla.WidthPercentage = 100
        tabla.SetWidths(anchoCeldas)
        tabla.AddCell("No.")
        tabla.AddCell("Error")
        tabla.AddCell("Tipo")
        tabla.AddCell("Descripción")
        tabla.AddCell("Fila")
        tabla.AddCell("Columna")

        '---Recorrido de los ArrayList
        For i = 0 To errorReconocido.Count - 1
            tabla.AddCell(i + 1)
            tabla.AddCell(errorReconocido.Item(i))
            tabla.AddCell("Léxico")
            tabla.AddCell("Caracter Desconocido")
            tabla.AddCell(filaerror(i))
            tabla.AddCell(columnaError(i))
        Next
        '---
        pdf.Add(tabla)
        '---
        pdf.Close()
    End Sub

    Friend Shared Sub crearPDFUser(lexemaReconocido As ArrayList, interlineado As Integer, tamanoLetra As Double, dirArchivo As String, archivoSalida As String)
        Dim pdfSalida As New Document(PageSize.LETTER, 70, 70, 50, 50)
        Dim pdfEscritor As PdfWriter

        Dim errorEncontrado As Boolean = False
        Dim texto As Boolean = False

        Try
            pdfEscritor = PdfWriter.GetInstance(pdfSalida, New FileStream(dirArchivo & archivoSalida, FileMode.Create, FileAccess.Write, FileShare.None))
        Catch ex As Exception
            pdfEscritor = PdfWriter.GetInstance(pdfSalida, New FileStream(My.Computer.FileSystem.SpecialDirectories.Desktop & "/" & archivoSalida, FileMode.Create, FileAccess.Write, FileShare.None))
        End Try

        pdfSalida.Open()

        Dim indice As Integer = 0

        While indice < lexemaReconocido.Count

            If (lexemaReconocido(indice).ToString.ToUpper = "TEXTO") Then
                texto = True
            ElseIf (texto = True And lexemaReconocido(indice).ToString = "}") Then
                texto = False
            End If


            If texto = True Then
                '---Texto subrayado o en negrita
                If (lexemaReconocido(indice).ToString.ToUpper = "[".ToUpper) Then
                    If lexemaReconocido(indice + 1).ToString.ToUpper = "+" Then
                        'Escribo texto en negrita
                        Dim linea As Paragraph = New Paragraph()
                        linea.Font = FontFactory.GetFont("Arial", Convert.ToInt32(tamanoLetra))
                        linea.Font.SetStyle(Font.BOLD)
                        linea.SetLeading(0, interlineado + 0.5)
                        linea.SpacingAfter = interlineado + 0.5
                        linea.Add(lexemaReconocido(indice + 2).ToString)

                        'indice += 2
                        pdfSalida.Add(linea)
                    ElseIf lexemaReconocido(indice + 1).ToString.ToUpper = "*" Then
                        'Escribo texto en raya
                        'indice += 1
                        Dim linea As Paragraph = New Paragraph
                        linea.Font = FontFactory.GetFont("Arial", Convert.ToInt32(tamanoLetra))
                        linea.Font.SetStyle(Font.STRIKETHRU)
                        linea.SetLeading(0, interlineado + 0.5)
                        linea.SpacingAfter = interlineado + 0.5
                        linea.Add(lexemaReconocido(indice + 2))
                        'indice += 2
                        pdfSalida.Add(linea)
                    End If
                End If
                '---Lista Enumerada
                If lexemaReconocido(indice).ToString.ToUpper = "Numeros".ToUpper Then
                    Dim numerado As String = ""
                    Dim palabras() As String

                    If lexemaReconocido(indice + 1).ToString.ToUpper = "(" Then
                        Do Until (lexemaReconocido(indice + 1).ToString.ToUpper = ")")
                            indice += 1
                            numerado = numerado + lexemaReconocido(indice + 1)
                        Loop

                        numerado = numerado.Replace(Chr(34), "")
                        numerado = numerado.Replace(Chr(34), "")
                        numerado = numerado.Replace(")", "")
                        palabras = numerado.Split(",")
                    End If
                    For i = 1 To palabras.Length()
                        Dim linea As Paragraph = New Paragraph
                        linea.Font = FontFactory.GetFont("Arial", Convert.ToInt32(tamanoLetra))
                        linea.Add(i & ". " & palabras(i - 1))
                        linea.SetLeading(0, interlineado + 1)
                        linea.SpacingAfter = interlineado + 1
                        pdfSalida.Add(linea)
                    Next
                    indice += 1
                End If
                '---Texto en Comilla
                If lexemaReconocido(indice).ToString.First = Chr(34) Then
                    Dim linea As Paragraph = New Paragraph
                    linea.Font = FontFactory.GetFont("Arial", Convert.ToInt32(tamanoLetra))

                    linea.Add(lexemaReconocido(indice).ToString.Replace(Chr(34), ""))
                    linea.SetLeading(0, interlineado + 1)
                    linea.SpacingAfter = interlineado + 1
                    pdfSalida.Add(linea)
                End If
                '---Linea en Blanco
                If lexemaReconocido(indice).ToString.ToUpper = "linea_en_blanco".ToUpper Then
                    pdfSalida.Add(Chunk.NEWLINE)
                End If
            End If

            '---Imagen y lo envio directamente...
            If (lexemaReconocido(indice).ToString.ToUpper = "Imagen".ToUpper) Then
                Dim dir_x_y As String = ""
                Dim DXY As String()
                If lexemaReconocido(indice + 1).ToString = "(" Then

                    Do Until lexemaReconocido(indice + 1).ToString = ")"
                        indice += 1
                        dir_x_y = dir_x_y + lexemaReconocido(indice + 1)
                    Loop
                    dir_x_y = dir_x_y.Replace(Chr(34), "")
                    dir_x_y = dir_x_y.Replace(Chr(34), "")
                    dir_x_y = dir_x_y.Replace(")", "")
                    DXY = dir_x_y.Split(",")
                End If
                Try
                    Dim imagenPDF As Image = Image.GetInstance(DXY(0))
                    imagenPDF.ScaleToFit(Convert.ToInt32(DXY(1)), Convert.ToInt32(DXY(2)))
                    pdfSalida.Add(imagenPDF)
                Catch ex As Exception
                    pdfSalida.Add(New Paragraph("La dirección no se ha encontrado..."))
                End Try
            End If

            indice += 1
        End While

        pdfSalida.Close()
    End Sub
End Class
