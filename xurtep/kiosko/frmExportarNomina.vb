﻿Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions

Public Class frmExportarNomina
    Dim sheetIndex As Integer = -1
    Dim SQL As String
    Dim contacolumna As Integer
    Public gIdFactura As String
    Private Sub tsbNuevo_Click(sender As Object, e As EventArgs) Handles tsbNuevo.Click
        tsbNuevo.Enabled = False
        tsbImportar.Enabled = True
        tsbImportar_Click(sender, e)
    End Sub

    Private Sub tsbImportar_Click(sender As Object, e As EventArgs) Handles tsbImportar.Click
        Dim dialogo As New OpenFileDialog
        lblRuta.Text = ""
        With dialogo
            .Title = "Búsqueda de archivos de saldos."
            .Filter = "Hoja de cálculo de excel (xlsx)|*.xlsx;"
            .CheckFileExists = True
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                lblRuta.Text = .FileName
            End If
        End With
        tsbProcesar.Enabled = lblRuta.Text.Length > 0
        If tsbProcesar.Enabled Then
            tsbProcesar_Click(sender, e)
        End If


    End Sub

    Private Sub tsbProcesar_Click(sender As Object, e As EventArgs) Handles tsbProcesar.Click
        lsvLista.Items.Clear()
        lsvLista.Columns.Clear()
        lsvLista.Clear()

        pnlCatalogo.Enabled = False
        tsbGuardar.Enabled = False
        tsbCancelar.Enabled = False
        lsvLista.Visible = False
        tsbImportar.Enabled = False
        Me.cmdCerrar.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        Application.DoEvents()

        Try
            If File.Exists(lblRuta.Text) Then
                Dim Archivo As String = lblRuta.Text
                Dim Hoja As String


                Dim book As New ClosedXML.Excel.XLWorkbook(Archivo)
                If book.Worksheets.Count >= 1 Then
                    sheetIndex = 1
                    If book.Worksheets.Count > 1 Then
                        Dim Forma As New frmHojasNomina
                        Dim Hojas As String = ""
                        For i As Integer = 0 To book.Worksheets.Count - 1
                            Hojas &= book.Worksheets(i).Name & IIf(i < (book.Worksheets.Count - 1), "|", "")
                        Next
                        Forma.Hojas = Hojas
                        If Forma.ShowDialog = Windows.Forms.DialogResult.OK Then
                            sheetIndex = Forma.selectedIndex + 1
                        Else
                            Exit Sub
                        End If
                    End If
                    Hoja = book.Worksheet(sheetIndex).Name
                    Dim sheet As IXLWorksheet = book.Worksheet(sheetIndex)

                    Dim colIni As Integer = sheet.FirstColumnUsed().ColumnNumber()
                    Dim colFin As Integer = sheet.LastColumnUsed().ColumnNumber()
                    Dim Columna As String
                    Dim numerocolumna As Integer = 1


                    lsvLista.Columns.Add("#")
                    For c As Integer = colIni To colFin

                        lsvLista.Columns.Add(sheet.Cell(1, c).Value)
                        'lsvLista.Columns.Add(numerocolumna)
                        'numerocolumna = numerocolumna + 1

                    Next

                    'lsvLista.Columns.Add("conciliacion")
                    'lsvLista.Columns.Add("color")

                    lsvLista.Columns(1).Name = "Fecha"
                    lsvLista.Columns(1).Width = 90
                    lsvLista.Columns(2).Width = 100
                    lsvLista.Columns(3).Width = 400
                    lsvLista.Columns(4).Width = 100
                    lsvLista.Columns(4).TextAlign = 1
                    lsvLista.Columns(5).Width = 400
                    lsvLista.Columns(6).Width = 100
                    lsvLista.Columns(6).TextAlign = 1
                    lsvLista.Columns(7).Width = 400
                    lsvLista.Columns(8).Width = 400
                    'lsvLista.Columns(8).TextAlign = 1
                    lsvLista.Columns(9).Width = 400
                    'lsvLista.Columns(10).Width = 400
                    'lsvLista.Columns(11).Width = 400
                    'lsvLista.Columns(12).Width = 400
                    'lsvLista.Columns(13).Width = 400


                    Dim Filas As Long = sheet.RowsUsed().Count()
                    For f As Integer = 2 To Filas
                        Dim item As ListViewItem = lsvLista.Items.Add((f - 1).ToString)

                        For c As Integer = colIni To colFin
                            Try

                                Dim Valor As String = ""
                                If (sheet.Cell(f, c).ValueCached Is Nothing) Then
                                    Valor = sheet.Cell(f, c).Value.ToString()
                                Else
                                    Valor = sheet.Cell(f, c).ValueCached.ToString()
                                End If
                                Valor = Valor.Trim()
                                item.SubItems.Add(Valor)


                                'If f = 6 And c >= 12 Then

                                '    'If Valor <> "" AndAlso InStr(Valor, "-") > 0 Then
                                '    '    Dim sValores() As String = Valor.Split("-")
                                '    '    Select Case sValores(0).ToUpper()
                                '    '        Case "P"
                                '    '            item.SubItems(item.SubItems.Count - 1).Tag = "1" 'Percepción
                                '    '        Case "D"
                                '    '            item.SubItems(item.SubItems.Count - 1).Tag = "2" 'Deducción
                                '    '        Case "I"
                                '    '            item.SubItems(item.SubItems.Count - 1).Tag = "3" 'Incapacidad
                                '    '    End Select
                                '    '    Valor = sValores(1).Trim()
                                '    'End If
                                '    item.SubItems(item.SubItems.Count - 1).Text = Valor
                                'End If



                            Catch ex As Exception

                            End Try

                        Next
                    Next

                    book.Dispose()
                    book = Nothing
                    GC.Collect()
                    'If lsvNominaFile.Items.Count >= 9 Then
                    '    If chkTipo.Checked Then
                    '        ProcesarNomina()
                    '    Else
                    '        ProcesarNomina1()
                    '    End If

                    'End If
                    pnlCatalogo.Enabled = True
                    If lsvLista.Items.Count = 0 Then
                        MessageBox.Show("El catálogo no puso ser importado o no contiene registros." & vbCrLf & "¿Por favor verifique?", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        MessageBox.Show("Se han encontrado " & FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        tsbGuardar.Enabled = True
                        tsbCancelar.Enabled = True
                        lblRuta.Text = FormatNumber(lsvLista.Items.Count, 0) & " registros en el archivo."
                        Me.Enabled = True
                        Me.cmdCerrar.Enabled = True
                        Me.Cursor = Cursors.Default
                        tsbImportar.Enabled = True
                        lsvLista.Visible = True
                    End If




                ElseIf book.Worksheets.Count = 0 Then
                    MessageBox.Show("El archivo no contiene hojas.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("El archivo ya no se encuentra en la ruta indicada.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Function getColumnName(index As Single) As String
        Dim numletter As Single = 26
        Dim sGrupo As Single = index / numletter
        Dim Modulo As Single = sGrupo - Math.Truncate(sGrupo)

        If Modulo = 0 Then Modulo = 1
        Dim Grupo As Integer = sGrupo - Modulo

        Dim Indice As Integer = index - (Grupo * numletter)
        Dim ColumnName As String = ""

        If Grupo > 0 Then
            ColumnName = Chr(64 + Grupo)
        End If
        ColumnName &= Chr(64 + Indice)
        Return ColumnName

    End Function

    Private Sub tsbCancelar_Click(sender As Object, e As EventArgs) Handles tsbCancelar.Click
        lsvLista.Items.Clear()
        lsvLista.Clear()
        pnlCatalogo.Enabled = False
        tsbGuardar.Enabled = False
        tsbCancelar.Enabled = False

        tsbNuevo.Enabled = True

    End Sub

    Private Sub tsbGuardar_Click(sender As Object, e As EventArgs) Handles tsbGuardar.Click
        Dim SQL As String, nombresistema As String = ""
        Try
            If lsvLista.CheckedItems.Count > 0 Then
                Dim mensaje As String


                pnlProgreso.Visible = True
                pnlCatalogo.Enabled = False
                Application.DoEvents()


                Dim IdProducto As Long
                Dim i As Integer = 0
                Dim conta As Integer = 0
                Dim bSubir As Boolean
                Dim Ruta As String
                pgbProgreso.Minimum = 0
                pgbProgreso.Value = 0
                pgbProgreso.Maximum = lsvLista.CheckedItems.Count


                'Dim fila As New DataRow


                bSubir = True

                For Each producto As ListViewItem In lsvLista.CheckedItems
                    'validar si existe o el archivo en el servidor
                    'Ruta = "\\pagina-pc\pagosnxurtep\" & Trim(producto.SubItems(5).Text) & ".pdf"

                    Dim Archivo As System.IO.FileInfo

                    If Trim(producto.SubItems(5).Text) <> "" Then



                        Ruta = "\\pagina-pc\pagosnxurtep\" & Trim(producto.SubItems(5).Text) & ".pdf"
                        Archivo = New System.IO.FileInfo(Ruta)
                        If (Archivo.Exists) Then

                        Else



                            bSubir = False
                            MessageBox.Show("Error: el nombre del archivo de pago SA no coincide con el almacenado en el servidor: Trabajador:" & Trim(producto.SubItems(3).Text) & ". La validación concluira en ese registro y no se subira ningun dato ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit For
                        End If
                    End If



                    If Trim(producto.SubItems(7).Text) <> "" Then
                        Ruta = "\\pagina-pc\pagosnxurtep\" & Trim(producto.SubItems(7).Text) & ".pdf"

                        Archivo = New System.IO.FileInfo(Ruta)
                        If (Archivo.Exists) Then

                        Else



                            bSubir = False
                            MessageBox.Show("Error: el nombre del archivo de pago Asimilados no coincide con el almacenado en el servidor: Trabajador:" & Trim(producto.SubItems(3).Text) & ". La validación concluira en ese registro y no se subira ningun dato ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit For
                        End If
                    End If




                    If Trim(producto.SubItems(8).Text) <> "" Then


                        Ruta = "\\pagina-pc\pagosnxurtep\" & Trim(producto.SubItems(8).Text) & ".pdf"

                        Archivo = New System.IO.FileInfo(Ruta)
                        If (Archivo.Exists) Then

                        Else



                            bSubir = False
                            MessageBox.Show("Error: el nombre del archivo timbrado Asimilados Pdf no coincide con el almacenado en el servidor: Trabajador:" & Trim(producto.SubItems(3).Text) & ". La validación concluira en ese registro y no se subira ningun dato ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit For
                        End If
                    End If


                    If Trim(producto.SubItems(9).Text) <> "" Then
                        Ruta = "\\pagina-pc\pagosnxurtep\" & Trim(producto.SubItems(9).Text) & ".xml"

                        Archivo = New System.IO.FileInfo(Ruta)
                        If (Archivo.Exists) Then

                        Else



                            bSubir = False
                            MessageBox.Show("Error: el nombre del archivo xml Asimilados no coincide con el almacenado en el servidor: Trabajador:" & Trim(producto.SubItems(3).Text) & ". La validación concluira en ese registro y no se subira ningun dato ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit For
                        End If
                    End If

                    


                Next

                If bSubir Then
                    For Each producto As ListViewItem In lsvLista.CheckedItems

                        SQL = "select * from usuarioK where codigo ='" & Trim(producto.SubItems(2).Text) & "'"
                        Dim rwUsuarioK As DataRow() = nConsulta(SQL)


                        If rwUsuarioK Is Nothing = False Then
                            'insertamos el pago
                            'Insertar nuevo
                            SQL = "EXEC setPagoInsertar   0," & rwUsuarioK(0)("iIdUsuarioK") & ",'" & Date.Parse(Trim(producto.SubItems(1).Text)).ToShortDateString() & "',1,"

                            If producto.SubItems(4).Text = "" Then
                                producto.SubItems(4).Text = "0"
                            End If
                            SQL &= Math.Round(Double.Parse(Trim(producto.SubItems(4).Text).Replace(",", "").Replace("$", "").ToString()), 2) & ","
                            SQL &= "'" & Trim(producto.SubItems(5).Text) & "',"

                            If producto.SubItems(6).Text = "" Then
                                producto.SubItems(6).Text = "0"
                            End If
                            SQL &= Math.Round(Double.Parse(Trim(producto.SubItems(6).Text).Replace(",", "").Replace("$", "").ToString()), 2) & ",'"

                            SQL &= Trim(producto.SubItems(7).Text) & "',"
                            SQL &= "0,'"
                            SQL &= IIf(chkNominaB.Checked = True, txtcarpeta.Text, "") & "','"
                            SQL &= Trim(producto.SubItems(8).Text) & "','"
                            SQL &= Trim(producto.SubItems(9).Text) & "',"
                            SQL &= "'',"
                            SQL &= "'',1"


                            If nExecute(SQL) = False Then

                                MessageBox.Show("Error en el registro con los siguiente datos: fecha:" & Trim(producto.SubItems(1).Text) & " trabajador:" & Trim(producto.SubItems(3).Text) & ". El proceso concluira en ese registro. ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Exit Sub

                            End If

                        Else
                            'insertamos tanto en usuario como en pago
                            'TransaccionKiosko = KIOSKOCONEXION.BeginTransaction

                            SQL = "EXEC setUsuarioKInsertar  0,'" & Trim(producto.SubItems(3).Text) & "','" & Trim(producto.SubItems(2).Text) & "','"
                            SQL &= Trim(producto.SubItems(2).Text) & "','" & Trim(producto.SubItems(2).Text) & "',0,0,1"


                            Dim idusuario As String
                            idusuario = ""
                            If Execute(SQL, idusuario) = False Then
                                MessageBox.Show("Error en el registro con los siguiente datos: fecha:" & Trim(producto.SubItems(1).Text) & " trabajador:" & Trim(producto.SubItems(3).Text) & ". El proceso concluira en ese registro. ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Exit Sub

                            End If

                            SQL = "EXEC setPagoInsertar   0," & idusuario & ",'" & Date.Parse(Trim(producto.SubItems(1).Text)).ToShortDateString() & "',1,"


                            If producto.SubItems(4).Text = "" Then
                                producto.SubItems(4).Text = "0"
                            End If
                            SQL &= Math.Round(Double.Parse(Trim(producto.SubItems(4).Text).Replace(",", "").Replace("$", "").ToString()), 2) & ","
                            SQL &= "'" & Trim(producto.SubItems(5).Text) & "',"

                            If producto.SubItems(6).Text = "" Then
                                producto.SubItems(6).Text = "0"
                            End If
                            SQL &= Math.Round(Double.Parse(Trim(producto.SubItems(6).Text).Replace(",", "").Replace("$", "").ToString()), 2) & ",'"
                            SQL &= Trim(producto.SubItems(7).Text) & "',"
                            SQL &= "0,'"
                            SQL &= IIf(chkNominaB.Checked = True, txtcarpeta.Text, "") & "','"
                            SQL &= Trim(producto.SubItems(8).Text) & "','"
                            SQL &= Trim(producto.SubItems(9).Text) & "',"
                            SQL &= "'',"
                            SQL &= "'',1"

                            If nExecute(SQL) = False Then

                                MessageBox.Show("Error en el registro con los siguiente datos: fecha:" & Trim(producto.SubItems(1).Text) & " trabajador:" & Trim(producto.SubItems(3).Text) & ". El proceso concluira en ese registro. ", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                Exit Sub

                            End If

                        End If


                        pgbProgreso.Value += 1
                        Application.DoEvents()

                    Next

                End If

                MessageBox.Show("Proceso terminado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                tsbCancelar_Click(sender, e)
                pnlProgreso.Visible = False



            Else

                MessageBox.Show("Por favor seleccione al menos una registro para importar.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            pnlCatalogo.Enabled = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    Private Sub chkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged
        For Each item As ListViewItem In lsvLista.Items
            item.Checked = chkAll.Checked
        Next
        chkAll.Text = IIf(chkAll.Checked, "Desmarcar todos", "Marcar todos")
    End Sub

    Private Sub cmdCerrar_Click(sender As Object, e As EventArgs) Handles cmdCerrar.Click
        Me.Close()
    End Sub

    Private Sub frmExportarNomina_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Cambio
        'dos
    End Sub
End Class