Public Class frmPrincipal
    Dim SQL As String
    Private Sub frmPrincipal_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If sender Is Nothing = False Then
            Dim Forma As Windows.Forms.Form = CType(sender, Windows.Forms.Form)
            Dim Nombre As String = Forma.Name
            If Nombre = "frmFacturar" Then
                chkCFDI.Visible = False
                AjustarBarra()
            ElseIf Nombre = "frmFacturaCBB" Then
                chkCBB.Visible = False
                AjustarBarra()
            End If
        End If
    End Sub

    Private Sub AjustarBarra()
        chkCFDI.Left = lblUsuario.Left + lblUsuario.Width + 10
        chkCBB.Left = lblUsuario.Left + lblUsuario.Width + IIf(chkCFDI.Visible, chkCFDI.Width + 5, 0) + 10
    End Sub

    Private Sub frmPrincipal_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("¿Desea salir del sistema?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub frmPrincipal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        lblUsuario.Text = Usuario.Nombre
        clsConfiguracion.Actualizar()
    End Sub

    Private Sub CatálogoDeClientesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        'Dim Forma As New Vendedores
        'Forma.ShowDialog()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim PT As Point = Me.PointToScreen(CheckBox1.Location)

        If CheckBox1.Checked Then
            MenuInicio.Show(CheckBox1.Location.X, (CheckBox1.Location.Y + pnlBar.Location.Y) - ((CheckBox1.Height - 4) * MenuInicio.Items.Count))
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub frmPrincipal_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged
        If sender Is Nothing = False Then
            Dim Forma As Windows.Forms.Form = CType(sender, Windows.Forms.Form)
            Dim Nombre As String = Forma.Name
            If Nombre = "frmFacturar" Then
                'RemoveHandler chkCFDI.CheckedChanged, AddressOf chkCFDI_CheckedChanged
                chkCFDI.Checked = Forma.WindowState <> FormWindowState.Minimized
                'frmFacturacionCFDI.Visible = frmFacturacionCFDI.WindowState <> FormWindowState.Minimized
                'AddHandler chkCFDI.CheckedChanged, AddressOf chkCFDI_CheckedChanged
            ElseIf Nombre = "frmFacturaCBB" Then
                'RemoveHandler chkCBB.CheckedChanged, AddressOf chkCBB_CheckedChanged
                chkCBB.Checked = Forma.WindowState <> FormWindowState.Minimized
                'frmFacturacionCBB.Visible = frmFacturacionCBB.WindowState <> FormWindowState.Minimized
                'AddHandler chkCBB.CheckedChanged, AddressOf chkCBB_CheckedChanged
            End If
        End If
    End Sub

    Private Sub lsvPanel_ItemActivate(sender As Object, e As System.EventArgs) Handles lsvPanel.ItemActivate
        Try
            If lsvPanel.SelectedItems.Count <= 0 Then
                Exit Sub
            End If
            Select Case lsvPanel.SelectedItems(0).Text
                Case "Exportar nomina"
                    Try

                        Dim forma As New frmExportarNomina
                        forma.ShowDialog()


                    Catch ex As Exception
                    End Try
                Case "Exportar anticipos nomina"
                    Try

                        Dim forma As New frmExportarAnticipo
                        forma.ShowDialog()


                    Catch ex As Exception
                    End Try
                Case "Exportar gastos viaje"
                    Try

                        Dim forma As New frmExportarGastosViaje
                        forma.ShowDialog()


                    Catch ex As Exception
                    End Try

            End Select

        Catch ex As Exception
            ShowError(ex, Me.Text)
        End Try
    End Sub

    Private Sub lsvPanel_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lsvPanel.SelectedIndexChanged

    End Sub

    Private Sub lsvPanel_SizeChanged(sender As Object, e As System.EventArgs) Handles lsvPanel.SizeChanged
        Dim sRuta As String
        sRuta = System.IO.Path.GetTempPath
        Try
            Me.lsvPanel.BackgroundImage = Me.PictureBox1.Image.GetThumbnailImage(Me.lsvPanel.ClientSize.Width, Me.lsvPanel.ClientSize.Height, Nothing, Nothing)
            Me.BackgroundImage = Me.PictureBox1.Image.GetThumbnailImage(Me.lsvPanel.ClientSize.Width, Me.lsvPanel.ClientSize.Height, Nothing, Nothing)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub pnlBar_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles pnlBar.Paint
        Degradado(e, sender, Color.White, Color.Gray, Drawing2D.LinearGradientMode.Vertical)
    End Sub

    Private Sub lblUsuario_SizeChanged(sender As Object, e As System.EventArgs) Handles lblUsuario.SizeChanged
        AjustarBarra()
    End Sub

    Private Sub mnuSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSalir.Click
        Me.Close()
    End Sub


    Private Sub CatalogosToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)


    End Sub


End Class
