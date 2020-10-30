Imports System.Net.Mail

Public Class Form1

    Private Property rutaArchivo As String = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rutaArchivo = "C:\Users\alumno\Desktop\Nuevo documento de texto.txt"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If IsNothing(rutaArchivo) Then Throw New Exception("Debe seleccionar una ruta valida")
        MsgBox(rutaArchivo)

        Dim listaPersonas As New List(Of Persona)
        Dim datosNoCargados As New List(Of String)

        Using archivo As New Microsoft.VisualBasic.FileIO.TextFieldParser(rutaArchivo)
            archivo.TextFieldType = FileIO.FieldType.Delimited
            archivo.SetDelimiters(":")
            Dim currentRow As String()

            While Not archivo.EndOfData
                Try
                    currentRow = archivo.ReadFields()
                    Dim linea As String() = currentRow.ToArray()

                    If linea.Length < 2 Then
                        'MsgBox("La linea con el dato " & linea(0) & " no se cargarà")
                        Dim lineaError As String = linea(0) & ":" & linea(1) & ":" & linea(2)
                        datosNoCargados.Add(lineaError)

                    Else
                        Dim persona As New Persona
                        persona.nombre = linea(0)
                        persona.dato1 = Integer.Parse(linea(1))
                        persona.dato2 = Integer.Parse(linea(2))

                        listaPersonas.Add(persona)
                    End If

                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox(ex.Message & vbCrLf & ex.ToString & vbCrLf & ex.StackTrace, vbCritical, "error al leer una lìnea")
                Catch ex As Exception
                    MsgBox(ex.Message & vbCrLf & ex.ToString & vbCrLf & ex.StackTrace, vbCritical, "Error")
                End Try

            End While

        End Using
        For Each persona As Persona In listaPersonas
            MsgBox("Nombre persona: " & persona.nombre & vbCrLf & "Dato 1: " & persona.dato1 & vbCrLf & "Dato 2: " & persona.dato2, vbInformation, "Datos cargados corectamente")
        Next

        If datosNoCargados.Count > 0 Then
            MsgBox("Los siguientes datos no se cargaron corectamente:")
            For Each cadena As String In datosNoCargados
                MsgBox(cadena, vbCritical, "Linea no cargada")
            Next
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' se muestra el OpenFileDialog
        ofdAbrirAArchivo.ShowDialog()
        ' se obtiene el String de la ruta
        rutaArchivo = ofdAbrirAArchivo.FileName
        TextBox1.Text = rutaArchivo
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

End Class
