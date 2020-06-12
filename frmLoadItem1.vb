Imports System.Data.OleDb

Public Class frmLoadItem1

    Private Sub LoadItem()
        Dim totalstock As Integer
        Dim C As Integer
        Try
            sqL = "SELECT ItemNo, itemCode, iDescription, StocksOnHand, DataExpira FROM Item Order By iDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4))
                C = dr(0)
                C = 1
                totalstock += C
            Loop
            lbltt1.Text = totalstock
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub Search()
        Dim totalstock As Integer
        Dim C As Integer
        Try
            sqL = "SELECT ItemNo, itemCode, iDescription, StocksOnHand, DataExpira FROM Item WHERE iDescription LIKE '" & TextBox1.Text & "%' Order By iDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4))
                C = dr(0)
                C = 1
                totalstock += C
            Loop
            lbltt1.Text = totalstock
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub frmLoadItem1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadItem()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Search()
    End Sub

    Private Sub dgw_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgw.CellDoubleClick
        If frmItem.search = True Then
            frmItem.txtItemNo.Text = dgw.CurrentRow.Cells(0).Value.ToString()
            frmItem.txtItemCode.Text = dgw.CurrentRow.Cells(1).Value.ToString()
            frmItem.txtDescription.Text = dgw.CurrentRow.Cells(2).Value.ToString()
            frmItem.txtQuantity.Text = dgw.CurrentRow.Cells(3).Value.ToString()
            frmItem.txtDataExpira.Text = dgw.CurrentRow.Cells(4).Value.ToString()
            frmItem.search = False
            Me.Close()
        End If



    End Sub

    Private Sub dgw_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgw.CellContentClick

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class