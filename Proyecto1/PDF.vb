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

    Friend Shared Sub crearPDFUser(lexemaReconocido As ArrayList, interlineado As Integer, tamanoLetra As Double, dirArchivo As String)
        Dim pdfSalida As New Document(PageSize.LETTER, 20, 20, 10, 10)
        Dim pdfEscritor As PdfWriter
        Try
            pdfEscritor = PdfWriter.GetInstance(pdfSalida, New FileStream(dirArchivo, FileMode.Create, FileAccess.Write, FileShare.None))
        Catch ex As Exception
        Finally
            pdfEscritor = PdfWriter.GetInstance(pdfSalida, New FileStream("C:\", FileMode.Create, FileAccess.Write, FileShare.None))
        End Try

        pdfSalida.Open()

        Dim indice As Integer = 0

        While indice < lexemaReconocido.Count
            '---Texto subrayado o en negrita
            If (lexemaReconocido(indice).ToString.ToUpper = "[".ToUpper) Then
                If lexemaReconocido(indice + 1).ToString.ToUpper = "+" Then
                    'Escribo texto en negrita

                    indice += 2
                ElseIf lexemaReconocido(indice + 1).ToString.ToUpper = "*" Then
                    'Escribo texto en raya

                    indice += 2
                End If
            Else
                indice += 1
            End If
            '---Lista Enumerada
            If lexemaReconocido(indice).ToString.ToUpper = "Numeros".ToUpper Then
                If lexemaReconocido(indice + 1).ToString.ToUpper = "(" Then

                End If
            End If
            '---
        End While
    End Sub
End Class
