
Imports System.Data.OleDb
Public Class frmReportStocksInPrint
    Dim fromMonth As String
    Dim toMonth As String
    Dim days As Integer
    Dim totalAmount As Double
    Dim y As Integer

    Private Sub LoadStocksIn()
        Dim totalQuantity As Integer
        Dim totalStocks As Integer
        If frmReportStocksIn.chkDaily.CheckState = CheckState.Checked Then
            sqL = "SELECT SIDate, IDescription, ItemQuantity, CurrentStocks FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND DAY(SIDate) =" & Format(frmReportStocksIn.dtpTo.Value, "dd") & " Order By SIDate, IDescription"
        End If
        If frmReportStocksIn.chkMonthly.CheckState = CheckState.Checked Then
            sqL = "SELECT SIDate, IDescription, ItemQuantity, CurrentStocks FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND MONTH(SIDate) =" & Format(frmReportStocksIn.dtpTo.Value, "MM") & " Order By SIDate, IDescription"
        End If
        Try
            'sqL = "SELECT SIDate, IDescription, ItemQuantity, StocksOnHand FROM Item as I, StocksIn as S WHERE I.ItemNo = S.ItemNo AND SIDate >= #" & frmReportStocksIn.dtpFrom.Text & "# AND SIDate <=#" & frmReportStocksIn.dtpTo.Text & "# Order By SIDate, IDescription"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totalQuantity = 0
            totalStocks = 0
            y = 0
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3))
                totalQuantity += dr(2)
                y += 19

            Loop
            lblStocksin.Text = totalQuantity
            Me.Height = Me.Height + y
            Me.Panel2.Height = Me.Panel2.Height + y
            Panel4.Location = New Point(8, 227 + y)
            Panel1.Location = New Point(7, 185 + y)
            dgw.Height = dgw.Height + y
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub



    Private Sub frmReportStocksInPrint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblDate.Text = Date.Now.ToString("dd/MM/yyyy")
        'Dim d1 As Date = Format(frmReportStocksIn.dtpFrom.Value, "Short Date")
        Dim d2 As Date = Format(frmReportStocksIn.dtpTo.Value, "Short Date")

        fromMonth = frmReportStocksIn.chkMonthly.CheckState = CheckState.Checked
        toMonth = frmReportStocksIn.dtpTo.Value.ToString("MMMM")

        'days = DateDiff(DateInterval.Day, d1, d2)
        If frmReportStocksIn.chkDaily.CheckState = CheckState.Checked Then
            'If days <= 7 Then
            lblreport.Text = "Inventário Diário"
        Else
            lblreport.Text = "Inventário Mensal"
        End If

        If fromMonth = "January" And toMonth = "March" Then
            lblreport.Text = "Inventário do 1o trimestre do Ano"
        ElseIf fromMonth = "April" And toMonth = "June" Then
            lblreport.Text = "Inventário do 2o trimestre do Ano"
        ElseIf fromMonth = "July" And toMonth = "September" Then
            lblreport.Text = "Inventário do 3o trimestre do Ano"
        ElseIf fromMonth = "October" And toMonth = "December" Then
            lblreport.Text = "Inventário do 4o trimestre do Ano"
        ElseIf fromMonth = "January" And toMonth = "December" Then
            lblreport.Text = "Inventário do Ano de " & frmReportSoldItems.dtpTo.Value.ToString("yyyy")

        End If
        Company()
        LoadStocksIn()
        PrintDialog1.Document = Me.PrintDocument1

        Dim ButtonPressed As DialogResult = PrintDialog1.ShowDialog()
        If (ButtonPressed = DialogResult.OK) Then
            PrintDocument1.Print()
        End If
        Me.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim bm As New Bitmap(Me.Panel2.Width, Me.Panel2.Height)

        Panel2.DrawToBitmap(bm, New Rectangle(0, 0, Me.Panel2.Width, Me.Panel2.Height))

        e.Graphics.DrawImage(bm, 0, 0)
        Dim aPS As New PageSetupDialog
        aPS.Document = PrintDocument1
    End Sub

    Private Sub Company()
        Try
            sqL = "SELECT HotelName  from Company"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            cmd.Parameters.AddWithValue("@d1", lblCompany.Text)
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                lblCompany.Text = dr.GetValue(0)

            End If
            If (dr IsNot Nothing) Then
                dr.Close()
            End If
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class