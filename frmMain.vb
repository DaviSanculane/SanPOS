Public Class frmMain

    Private Sub ProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        frmItem.ShowDialog()
    End Sub

    Private Sub StaffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StaffToolStripMenuItem.Click
        frmStaff.ShowDialog()
    End Sub

    Private Sub CustomersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomersToolStripMenuItem.Click
        frmCustomer.ShowDialog()
    End Sub

    Private Sub PointOfSaleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PointOfSaleToolStripMenuItem.Click
        frmPOS.ShowDialog()
    End Sub

    Private Sub AvailableStocksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvailableStocksToolStripMenuItem.Click
        frmAvailableStocks.ShowDialog()
    End Sub

    Private Sub SoldItemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmSoldItems.ShowDialog()
    End Sub

    Private Sub SoldItemsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SoldItemsToolStripMenuItem1.Click
        frmReportSoldItems.ShowDialog()
    End Sub

    Private Sub StocksInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StocksInToolStripMenuItem.Click
        frmReportStocksIn.ShowDialog()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        frmStaff.ShowDialog()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        frmItem.ShowDialog()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        frmCustomer.ShowDialog()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        frmPOS.ShowDialog()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        frmAvailableStocks.ShowDialog()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        frmReportSoldItems.ShowDialog()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'frmReportSoldItems.ShowDialog()
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        frmReportStocksIn.ShowDialog()
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblDate.Text = Date.Now.ToString("dddd") & " - " & Date.Now.ToString("MMM dd, yyyy")
        frmLogin.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lblTIme.Text = Format(Date.Now, "Long Time")
    End Sub

    Private Sub UsersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsersToolStripMenuItem.Click
        frmUsers.ShowDialog()
    End Sub

    Private Sub lblDate_Click(sender As Object, e As EventArgs) Handles lblDate.Click

    End Sub

    Private Sub ConfigurarEmpresaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfigurarEmpresaToolStripMenuItem.Click
        FrmCompany.ShowDialog()
    End Sub
    Sub LogOut()

        Me.Hide()
        frmLogin.Show()
        frmLogin.txtUser.Text = ""
        frmLogin.txtPwd.Text = ""
        frmLogin.txtUser.Focus()
    End Sub
    Private Sub RocarUsuárioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RocarUsuárioToolStripMenuItem.Click
        Try
            If MessageBox.Show("Tem certeza que desejas trocar usuário?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

                LogOut()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.davidsoft.webs.com")
    End Sub
End Class
