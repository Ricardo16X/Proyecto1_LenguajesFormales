Imports itextsharp.text
Imports itextsharp.text.pdf
Imports System.IO
Public Class PDF
    Friend Shared Sub crearPDF(Lexemas As ArrayList, tipodeToken As ArrayList, nomArchivo As String)

        'ruta del documento

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
        pdf.Add(New Paragraph("Archivo Fuente: " & nomArchivo))
        '---

        '---Tabla con los Lexemas recopilados en el analizador...
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)
        pdf.Add(Chunk.NEWLINE)

        Dim anchoCeldas() As Integer = {2, 10, 6, 2, 2}
        Dim tabla As New PdfPTable(5)
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
            tabla.AddCell("Tipo " & i)
            tabla.AddCell("Fila " & i)
            tabla.AddCell("Columna " & i)
        Next
        '---
        pdf.Add(tabla)
        '---
        pdf.Close()
    End Sub
End Class
