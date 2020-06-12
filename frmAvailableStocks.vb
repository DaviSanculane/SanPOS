

Imports System.Data.OleDb
Public Class frmAvailableStocks

    Private Sub LoadItems()
        Dim totalItems As Integer
        Dim totalstock As Integer
        Dim C As Integer
        Try
            sqL = "SELECT ItemNo, ItemCode, IDescription, ISize, UnitPrice, StocksOnHand, DataExpira FROM ITEM Where StocksOnHand > 0 Order By IDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4), dr(5), dr(6))
                totalItems += dr(5)
                C = dr(0)
                C = 1
                totalstock += C
            Loop
            lblTotalStocks.Text = totalItems
            lbltt.Text = totalstock
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub frmAvailableStocks_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadItems()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        frmPrintAvailableStocks.Show()
    End Sub

    Private Sub Label2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Search1()
        Dim totalItems As Integer
        Dim totalstock As Integer
        Dim C As Integer
        Try
            sqL = "SELECT ItemNo, ItemCode, iDescription, ISize, UnitPrice, StocksOnHand, DataExpira FROM Item WHERE iDescription LIKE '" & search.Text & "%' Order By iDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4), dr(5), dr(6))
                totalItems += dr(5)
                C = dr(0)
                C = 1
                totalstock += C
            Loop
            lblTotalStocks.Text = totalItems
            lbltt.Text = totalstock
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub search_TextChanged(sender As Object, e As EventArgs) Handles search.TextChanged
        Search1()
    End Sub
End Class