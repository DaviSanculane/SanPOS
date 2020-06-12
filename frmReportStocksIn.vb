

Imports System.Data.OleDb
Public Class frmReportStocksIn

    Private Sub LoadStocksIn()
        Dim totalQuantity As Integer
        Dim totalStocks As Integer

        If chkDaily.CheckState = CheckState.Checked Then
            sqL = "SELECT SIDate, IDescription, ItemQuantity, CurrentStocks FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND DAY(SIDate) =" & Format(dtpTo.Value, "dd") & " Order By SIDate, IDescription"
        End If
        If chkMonthly.CheckState = CheckState.Checked Then
            sqL = "SELECT SIDate, IDescription, ItemQuantity, CurrentStocks FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND MONTH(SIDate) =" & Format(dtpTo.Value, "MM") & " Order By SIDate, IDescription"
        End If
        Try
            'sqL = "SELECT SIDate, IDescription, ItemQuantity, CurrentStocks FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND SIDate >= #" & dtpFrom.Text & "# AND SIDate <=#" & dtpTo.Text & "# Order By SIDate, IDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totalQuantity = 0
            totalStocks = 0
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3))
                totalQuantity += dr(2)
                totalStocks += dr(2)
            Loop
            lblQuantity.Text = Format(totalQuantity, "#,##0")
            'lblTotal.Text = Format(totalStocks, "#,##0")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub frmReportStocksIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblQuantity.Text = ""
        'lblTotal.Text = ""
        dgw.Rows.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If chkDaily.CheckState = CheckState.Unchecked And chkMonthly.CheckState = CheckState.Unchecked Then
            MsgBox("Por favor selecione Diario ou Mensal", MsgBoxStyle.Critical, "Select")
            Exit Sub
        End If
        LoadStocksIn()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmReportStocksInPrint.Show()
    End Sub

    Private Sub chkDaily_CheckedChanged(sender As Object, e As EventArgs) Handles chkDaily.CheckedChanged
        If chkDaily.CheckState = CheckState.Checked Then
            chkMonthly.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub chkMonthly_CheckedChanged(sender As Object, e As EventArgs) Handles chkMonthly.CheckedChanged
        If chkMonthly.CheckState = CheckState.Checked Then
            chkDaily.CheckState = CheckState.Unchecked
        End If
    End Sub
End Class