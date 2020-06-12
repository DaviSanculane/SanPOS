Imports System.Data.OleDb

Public Class FrmCompany
    Sub Reset()

        txtTIN.Text = ""
        txtEmailID.Text = ""
        txtContactNo.Text = ""
        txtHotelName.Text = ""
        txtAddressLine1.Text = ""
        txtTicketFooterMessage.Text = ""
        txtStartBillNo.Text = 1
        txtStartBillNo.ReadOnly = False
        txtStartBillNo.Enabled = True
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        txtHotelName.Focus()
    End Sub
    Private Sub txtTicketFooterMessage_TextChanged(sender As Object, e As EventArgs) Handles txtTicketFooterMessage.TextChanged

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub txtStartBillNo_TextChanged(sender As Object, e As EventArgs) Handles txtStartBillNo.TextChanged

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtHotelName.Text = "" Then
            MessageBox.Show("Porfavor insira o nome da empresa", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtHotelName.Focus()
            Return
        End If
        If txtAddressLine1.Text = "" Then
            MessageBox.Show("Porfavor insira o endereço", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAddressLine1.Focus()
            Return
        End If

        If txtContactNo.Text = "" Then
            MessageBox.Show("Porfavor insira o contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Return
        End If

        If txtEmailID.Text = "" Then
            MessageBox.Show("Porfavor insira o Email", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtEmailID.Focus()
            Return
        End If

        Try

            sqL = "insert into Company( HotelName, AddressLine1, ContactNo, EmailID, NUit, StartBillNo ) VALUES ('" & txtHotelName.Text & "', '" & txtAddressLine1.Text & "', '" & txtContactNo.Text & "', '" & txtEmailID.Text & "', '" & txtTIN.Text & "', '" & txtStartBillNo.Text & "' )"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            cmd.Connection = conn
            'cmd.Parameters.AddWithValue("@d1", txtHotelName.Text)
            'cmd.Parameters.AddWithValue("@d2", txtAddressLine1.Text)
            'cmd.Parameters.AddWithValue("@d3", txtContactNo.Text)
            'cmd.Parameters.AddWithValue("@d4", txtEmailID.Text)
            'cmd.Parameters.AddWithValue("@d5", txtTIN.Text)
            'cmd.Parameters.AddWithValue("@d6", txtTicketFooterMessage.Text)
            'cmd.Parameters.AddWithValue("@d7", txtStartBillNo.Text)
            cmd.ExecuteNonQuery()
            conn.Close()
            'Dim st As String = "Empresa adicionada '" & txtHotelName.Text & "' info"
            'LogFunc(lblUser.Text, st)
            MessageBox.Show("Guardado com sucesso", "Empresa Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Getdata()
            Reset()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If txtHotelName.Text = "" Then
            MessageBox.Show("Porfavor insira o nome da empresa", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtHotelName.Focus()
            Return
        End If
        If txtAddressLine1.Text = "" Then
            MessageBox.Show("Porfavor insira o endereço", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAddressLine1.Focus()
            Return
        End If

        If txtContactNo.Text = "" Then
            MessageBox.Show("Porfavor insira o contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Return
        End If

        If txtEmailID.Text = "" Then
            MessageBox.Show("Porfavor insira o email id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtEmailID.Focus()
            Return
        End If

        Try


            sqL = "Update Company set HotelName = '" & txtHotelName.Text & "', AddressLine1 = '" & txtAddressLine1.Text & "', ContactNo = '" & txtContactNo.Text & "', EmailID = '" & txtEmailID.Text & "', Nuit = '" & txtTIN.Text & "'  where ID=" & txtID.Text & ""
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            cmd.Connection = conn
            'cmd.Parameters.AddWithValue("@d1", txtHotelName.Text)
            'cmd.Parameters.AddWithValue("@d2", txtAddressLine1.Text)
            'cmd.Parameters.AddWithValue("@d3", txtContactNo.Text)
            'cmd.Parameters.AddWithValue("@d4", txtEmailID.Text)
            'cmd.Parameters.AddWithValue("@d5", txtTIN.Text)
            'cmd.Parameters.AddWithValue("@d6", txtTicketFooterMessage.Text)
            cmd.ExecuteNonQuery()
            conn.Close()
            'sqL = "updated the restaurant '" & txtHotelName.Text & "' info"
            'LogFunc(lblUser.Text, sqL)
            MessageBox.Show("Actualizado com sucesso", "Empresa Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub Getdata()
        Try
            sqL = "SELECT ID, HotelName,  AddressLine1, ContactNo, EmailID, Nuit, StartBillNo from Company"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (dr.Read() = True)
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3), dr(4), dr(5), dr(6))
            End While
            conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Tem certeza que quer apagara esse record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtStartBillNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtStartBillNo.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Private Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0

            sqL = "delete from Company where ID=@d1"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)

            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "Empresa deletada '" & txtHotelName.Text & "' info"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Apagado com sucesso", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Getdata()
                Reset()
            Else
                MessageBox.Show("Registro não encontrado", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If conn.State = ConnectionState.Open Then
                conn.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub LogFunc(text As String, st As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub dgw_Paint(sender As Object, e As PaintEventArgs) Handles dgw.Paint

    End Sub

    Private Sub dgw_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub FrmCompany_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(Me.ActiveControl, True, True, True, False) 'for Select Next Control
        End If
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtID.Text = dr.Cells(0).Value.ToString()
                txtHotelName.Text = dr.Cells(1).Value.ToString()
                txtAddressLine1.Text = dr.Cells(2).Value.ToString()
                txtContactNo.Text = dr.Cells(3).Value.ToString()
                txtEmailID.Text = dr.Cells(4).Value.ToString()
                txtTIN.Text = dr.Cells(5).Value.ToString()
                'txtTicketFooterMessage.Text = dr.Cells(6).Value.ToString()
                txtStartBillNo.Text = dr.Cells(6).Value.ToString()
                btnUpdate.Enabled = True
                btnSave.Enabled = False
                btnDelete.Enabled = True
                txtStartBillNo.ReadOnly = True
                txtStartBillNo.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmCompany_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Getdata()
    End Sub
End Class