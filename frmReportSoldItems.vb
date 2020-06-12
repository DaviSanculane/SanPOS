
Imports System.Data.OleDb
Public Class frmReportSoldItems

    Private Sub LoadSoldItemsReport()
        Dim totalAmount As Double
        Dim totalQuantity As Integer

        If chkDaily.CheckState = CheckState.Checked Then
            sqL = "SELECT IDescription, UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND DAY(POSDate) =" & Format(dtpTo.Value, "dd") & " Group By IDescription, UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
        End If
        If chkMonthly.CheckState = CheckState.Checked Then
            sqL = "SELECT IDescription, UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND MONTH(POSDate) =" & Format(dtpTo.Value, "MM") & " Group By IDescription, UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
        End If


        Try
            'sqL = "SELECT IDescription, I.UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (I.UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND POSDate >= #" & dtpFrom.Text & "# AND POSDate <=#" & dtpTo.Text & "# Group By IDescription, I.UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()

            totalAmount = 0
            totalQuantity = 0
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4), dr(5))
                totalAmount += dr(5)
                totalQuantity += dr(2)
            Loop

            lblQuantity.Text = totalQuantity
            lblTotal.Text = Format(totalAmount, "#,##0.00")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If chkDaily.CheckState = CheckState.Unchecked And chkMonthly.CheckState = CheckState.Unchecked Then
            MsgBox("Por favor selecione Diario ou Mensal", MsgBoxStyle.Critical, "Select")
            Exit Sub
        End If
        LoadSoldItemsReport()
    End Sub

    Private Sub chkDaily_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDaily.CheckedChanged
        If chkDaily.CheckState = CheckState.Checked Then
            chkMonthly.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMonthly.CheckedChanged
        If chkMonthly.CheckState = CheckState.Checked Then
            chkDaily.CheckState = CheckState.Unchecked
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmReportSoldItemsPrint.Show()
    End Sub

    Private Sub frmReportSoldItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblQuantity.Text = ""
        lblTotal.Text = ""
        dgw.Rows.Clear()
    End Sub

    Private Sub txtSearchByDish_TextChanged(sender As Object, e As EventArgs) Handles txtSearchByDish.TextChanged
        Dim totalAmount As Double
        Dim totalQuantity As Integer

        If chkDaily.CheckState = CheckState.Checked Then
            sqL = "SELECT IDescription, UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND DAY(POSDate) =" & Format(dtpTo.Value, "dd") & " AND IDescription like '" & txtSearchByDish.Text & "%' Group By IDescription, UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
        End If
        If chkMonthly.CheckState = CheckState.Checked Then
            sqL = "SELECT IDescription, UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND MONTH(POSDate) =" & Format(dtpTo.Value, "MM") & " AND IDescription like '" & txtSearchByDish.Text & "%' Group By IDescription, UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
        End If
        Try
            'sqL = "SELECT IDescription, UnitPrice, SUM([Quantity]) as ItemQuantity, POSDate, POSTime, (UnitPrice * SUM([Quantity])) As TotalAmount FROM Item as I, POS as P, POSDetail as PD WHERE I.ItemNo = PD.ItemNo AND PD.InvoiceNo=P.InvoiceNo AND POSDate >= #" & dtpFrom.Text & "# AND POSDate <=#" & dtpTo.Text & "# AND IDescription like '" & txtSearchByDish.Text & "%' Group By IDescription, UnitPrice, [Quantity], POSDate, POSTime ORder By POSDate"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While (dr.Read() = True)
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4), dr(5))
                'dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(4))
                totalAmount += dr(5)
                totalQuantity += dr(2)
            Loop

            lblQuantity.Text = totalQuantity
            lblTotal.Text = Format(totalAmount, "#,##0.00")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub dtpFrom_ValueChanged(sender As Object, e As EventArgs)

    End Sub
End Class