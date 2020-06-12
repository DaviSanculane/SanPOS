Imports System.Data.OleDb
Public Class frmCustomer
    Dim adding As Boolean
    Dim updating As Boolean
    Public search As Boolean

    Private Sub GetCustomerNo()
        Try
            sqL = "SELECT CustomerNo FROM Customer Order By CustomerNo Desc"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtCustNo.Text = dr(0) + 1
            Else
                txtCustNo.Text = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub AddCustomer()
        Try
            sqL = "INSERT INTO Customer(Custname, Address, ContactNo) Values('" & txtName.Text & "', '" & txtAddress.Text & "', '" & txtContactNo.Text & "')"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Cliente adicionado", MsgBoxStyle.Information, "Add Customer")
            Else
                MsgBox("Falha ao adicionar cliente", MsgBoxStyle.Critical, "Add Customer")

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub UpdateCustomer()
        Try
            sqL = "Update Customer SET Custname ='" & txtName.Text & "', Address ='" & txtAddress.Text & "', ContactNo = '" & txtContactNo.Text & "' WHERE CustomerNo = " & txtCustNo.Text & ""
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Cliente actualizado", MsgBoxStyle.Information, "Update Customer")
            Else
                MsgBox("Falha ao actualizar cliente", MsgBoxStyle.Critical, "Update Customer")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub GetCustomerRecord()
        Try
            sqL = "SELECT CustName, Address, ContactNo FROM Customer WHERE CustomerNo = " & txtCustNo.Text & ""
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtName.Text = dr(0)
                txtAddress.Text = dr(1)
                txtContactNo.Text = dr(2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub ClearFields()
        txtCustNo.Text = ""
        txtName.Text = ""
        txtContactNo.Text = ""
        txtAddress.Text = ""
    End Sub

    Private Sub EnabledText()
        txtCustNo.Enabled = True
        txtName.Enabled = True
        txtAddress.Enabled = True
        txtContactNo.Enabled = True
    End Sub

    Private Sub DisabledText()
        txtCustNo.Enabled = False
        txtName.Enabled = False
        txtAddress.Enabled = False
        txtContactNo.Enabled = False
    End Sub


    Private Sub EnabledButton()
        btnAdd.Enabled = True
        btnUpdate.Enabled = True
        btnSearch.Enabled = True
        btnClose.Enabled = True

        btnSave.Enabled = False
        btnCancel.Enabled = False
    End Sub

    Private Sub DisabledButton()
        btnAdd.Enabled = False
        btnUpdate.Enabled = False
        btnSearch.Enabled = False
        btnClose.Enabled = False

        btnSave.Enabled = True
        btnCancel.Enabled = True
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        DisabledButton()
        ClearFields()
        EnabledText()
        GetCustomerNo()

        adding = True
        updating = False
        txtName.Focus()
        txtCustNo.Enabled = False
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If txtCustNo.Text = "" Then
            MsgBox("Por favor selecione o registro para actualizar", MsgBoxStyle.Critical, "Update Record")
            Exit Sub
        End If

        EnabledText()
        DisabledButton()

        adding = False
        updating = True
        txtName.Focus()
        txtCustNo.Enabled = False

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        search = True
        frmLoadCustomer.ShowDialog()
    End Sub

    Private Sub txtCustNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustNo.TextChanged
        If search = True Then
            GetCustomerRecord()
            search = False
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If adding = True Then
            AddCustomer()
            DisabledText()
            EnabledButton()
            ClearFields()
        Else
            UpdateCustomer()
            DisabledText()
            EnabledButton()
            ClearFields()
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        DisabledText()
        EnabledButton()
        ClearFields()
    End Sub

    Private Sub frmCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EnabledButton()
        DisabledText()
    End Sub

    Private Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            sqL = "delete from Customer where CustomerNo=@d1"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            cmd.Parameters.AddWithValue("@d1", Val(txtCustNo.Text))
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                'Dim st As String = "deleted the restaurant '" & txtHotelName.Text & "' info"
                MessageBox.Show("Cliente apagado com sucesso", "Resgistro", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else
                MessageBox.Show("Registro nao encontrado", "Desculpe", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If conn.State = ConnectionState.Open Then
                conn.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btApagar_Click(sender As Object, e As EventArgs) Handles btApagar.Click
        DeleteRecord()
        ClearFields()
    End Sub
End Class