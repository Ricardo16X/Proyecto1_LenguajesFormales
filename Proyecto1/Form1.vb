Imports System.IO

Public Class Form1
    'Variable para guardar la ruta del archivo
    Dim DireccionArchivo As String = ""
    'Información del Autor
    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        MessageBox.Show("Nombre: Ricardo Ismael Pérez Ajanel" & vbLf & "Carné: 201700524" & vbLf & vbLf & "Versión de Programa: 0.1")
    End Sub
    'Salir de la Aplicación
    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub
    'Abrir archivo con extensión .ACK
    Private Sub AbrirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbrirToolStripMenuItem.Click
        Dim openDialog As New OpenFileDialog
        openDialog.Filter = "Archivo .ack | *.ack"
        Try
            If (openDialog.ShowDialog) = DialogResult.OK Then
                DireccionArchivo = openDialog.FileName
                Me.CajaTexto.LoadFile(openDialog.FileName, RichTextBoxStreamType.PlainText)
            End If
        Catch ex As Exception
            MessageBox.Show("El archivo no se abrió correctamente!" & vbLf & "Intentelo Nuevamente")
        End Try
    End Sub
    'Guardar Archivo tanto si existe o no.
    Private Sub GuardarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuardarToolStripMenuItem.Click
        If DireccionArchivo.Length > 0 Then
            CajaTexto.SaveFile(DireccionArchivo, fileType:=RichTextBoxStreamType.PlainText)
        Else
            guardarArchivo(CajaTexto.Text)
        End If
    End Sub
    Private Sub guardarArchivo(Texto As String)
        Dim GuardarArchivo As New SaveFileDialog
        GuardarArchivo.Filter = "Archivo .ack | *.ack"
        Try
            If (GuardarArchivo.ShowDialog = DialogResult.OK) Then
                My.Computer.FileSystem.WriteAllText(GuardarArchivo.FileName, Texto, False)
                DireccionArchivo = GuardarArchivo.FileName
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    'Función para Analizar el Texto para poder crear el PDF con las instrucciones que aquí se escriban.
    Private Sub AnalizarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalizarToolStripMenuItem.Click
        If CajaTexto.Text.Length > 0 Then
            If DireccionArchivo.Length = 0 Then
                MessageBox.Show("Por favor, primero guarde el archivo.", "ACK - Generador PDF v0.1")
            Else
                Analizador.AnalisisLexico(CajaTexto.Text + "_", Path.GetFileName(DireccionArchivo))
            End If
        Else
            MessageBox.Show("No hay texto que analizar." & vbLf& & vbLf & "Por favor, cargue un archivo .ACK o escriba las respectivas instrucciones.")
        End If
    End Sub
End Class
