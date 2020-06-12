
Imports System.Data.OleDb
Public Class frmLogin


    Private Sub Login()
        Try
            sqL = "SELECT * FROM Users WHERE Username = '" & txtUser.Text & "' AND pwd = '" & txtPwd.Text & "'"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                frmMain.lblEmployeeNo.Text = dr("StaffID")
                frmMain.lblUser.Text = txtUser.Text
                frmMain.lblUserType.Text = dr("role")
                UserType.Text = dr("role")
                If (UserType.Text = "User") Then
                    Me.Hide()
                    frmMain.ManageToolStripMenuItem.Enabled = False
                    frmMain.ToolStripButton2.Enabled = False
                    frmMain.ToolStripButton1.Enabled = False
                    frmMain.ToolStripButton3.Enabled = False
                    'frmMain.lblUser.Text = txtUser.Text
                    'frmMain.lblUserType.Text = UserType.Text
                    frmMain.Show()
                End If

                If (UserType.Text = "Admin") Then
                    Me.Hide()
                    frmMain.ManageToolStripMenuItem.Enabled = True
                    frmMain.ToolStripButton2.Enabled = True
                    frmMain.ToolStripButton1.Enabled = True
                    frmMain.ToolStripButton3.Enabled = True
                    'frmMain.lblUser.Text = txtUser.Text
                    'frmMain.lblUserType.Text = UserType.Text
                    frmMain.Show()
                End If
                txtUser.Text = ""
                txtPwd.Text = ""
                Me.Close()
            Else
                MsgBox("Incorrect username or password!", MsgBoxStyle.Critical, "Login")
                txtPwd.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
            frmMain.Show()
        End Try
    End Sub


   
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Login()
        isAvailableQuantity()
        'isAvailableDate()
    End Sub

    Private Sub txtUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.GotFocus
        AcceptButton = btnLogin
    End Sub


    Private Sub txtPwd_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPwd.GotFocus
        AcceptButton = btnLogin
    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click

        If MsgBox("Are you sure you want to close?", MsgBoxStyle.YesNo, "Close Window") = MsgBoxResult.Yes Then
            End
        End If
        txtUser.Focus()
    End Sub
    Private Sub Login1()
        If Len(Trim(txtUser.Text)) = 0 Then
            MessageBox.Show("Please enter user Username/Porfavor insira Username", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUser.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPwd.Text)) = 0 Then
            MessageBox.Show("Please enter password/Porfavor insira a senha", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPwd.Focus()
            Exit Sub
        End If
        Try
            sqL = "SELECT RTRIM(Username),RTRIM(pwd),RTRIM(StaffID) FROM Users where Username = @d1 and pwd=@d2 and StaffID=@d3"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            cmd.Parameters.AddWithValue("@d1", txtUser.Text)
            cmd.Parameters.AddWithValue("@d2", txtPwd.Text)
            cmd.Parameters.AddWithValue("@d3", frmMain.lblEmployeeNo.Text)
            'frmMain.lblEmployeeNo.Text = dr("StaffID")
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                If txtPwd.Text.Trim = dr.GetValue(1).trim Then
                    sqL = "SELECT role FROM Users where Username=@d3 and pwd=@d4"
                    ConnDB()
                    cmd = New OleDbCommand(sqL, conn)
                    cmd.Parameters.AddWithValue("@d3", txtUser.Text)
                    cmd.Parameters.AddWithValue("@d4", txtPwd.Text)
                    dr = cmd.ExecuteReader()
                    If dr.Read() Then
                        UserType.Text = dr.GetValue(0).ToString.Trim
                    End If
                    If (dr IsNot Nothing) Then
                        dr.Close()
                    End If
                    If conn.State = ConnectionState.Open Then
                        conn.Close()
                    End If
                    If (UserType.Text = "User") Then
                        Me.Hide()
                        frmMain.ManageToolStripMenuItem.Enabled = False
                        frmMain.ToolStripButton2.Enabled = False
                        frmMain.ToolStripButton1.Enabled = False
                        frmMain.ToolStripButton3.Enabled = False
                        frmMain.lblUser.Text = txtUser.Text
                        frmMain.lblUserType.Text = UserType.Text
                        frmMain.Show()
                    End If

                    If (UserType.Text = "Admin") Then
                        Me.Hide()
                        frmMain.ManageToolStripMenuItem.Enabled = True
                        frmMain.ToolStripButton2.Enabled = True
                        frmMain.ToolStripButton1.Enabled = True
                        frmMain.ToolStripButton3.Enabled = True
                        frmMain.lblUser.Text = txtUser.Text
                        frmMain.lblUserType.Text = UserType.Text
                        frmMain.Show()
                    End If

                End If
            Else
                MsgBox("Login is Failed...Try again !", MsgBoxStyle.Critical, "Login Denied")
                txtUser.Text = ""
                txtPwd.Text = ""
                txtUser.Focus()
            End If
            conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function isAvailableQuantity() As Boolean
        Try
            sqL = "SELECT StocksOnHand FROM ITEM"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)


            If Val("StocksOnHand") <= 5 Then
                isAvailableQuantity = True
            Else
                MsgBox("Por favor, o Stock de alguns items estão prestes a terminar, verifique-os ", MsgBoxStyle.Information, "Stock Baixo")

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Function

    Private Function isAvailableDate() As Boolean
        Try
            sqL = "SELECT DataExpira FROM ITEM"
            ConnDB()
            cmd = New OleDbCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)


            If Val("DataExpira") <= FormatDateTime(Date.Now.ToString("dd-MM-yyyy")) Then
                isAvailableDate = True
            Else
                MsgBox("Por favor, a Data de alguns items estão prestes a expirar, verifique-os ", MsgBoxStyle.Information, "Stock Expirado")

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Function
End Class