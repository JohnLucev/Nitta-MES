Imports System
Imports System.Data.SqlClient


Public Class D3_Dry

    Dim AutomationTimer_Count As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AutomationTimer_Count = 0
        Timer1.Enabled = False
        AnimationTimer.Enabled = False


        Duct_Cover.BringToFront()
        Duct_Cover.Visible = True

        Access_Door_1.BringToFront()
        Access_Door_2.BringToFront()
        Access_Door_3.BringToFront()
        Access_Door_4.BringToFront()
        Access_Door_5.BringToFront()
        Access_Door_6.BringToFront()

        Access_Door_1.Visible = True
        Access_Door_2.Visible = True
        Access_Door_3.Visible = True
        Access_Door_4.Visible = True
        Access_Door_5.Visible = True
        Access_Door_6.Visible = True

        Hide_Hum_Controls.Visible = False
        Hide_Process_Air.Visible = False

        Arrow1A.Visible = False
        Arrow1B.Visible = False
        Arrow1C.Visible = False
        Arrow1D.Visible = False
        Arrow2A.Visible = False
        Arrow2B.Visible = False
        Arrow2C.Visible = False
        Arrow2D.Visible = False
        Arrow3A.Visible = False

        txtDevice_1.Text = ""
        txtDevice_1.Visible = False

        Z1a1_Blower.FillColor = Color.Red
        Z1a2_Blower.FillColor = Color.Red
        Z1B_Blower.FillColor = Color.Red
        Z1C1_Blower.FillColor = Color.Red
        Z1c2_Blower.FillColor = Color.Red
        Exhaust_Blower.FillColor = Color.Green

        txtZ1a1_Blower_on_off.Text = "0" 'False
        txtZ1a2_Blower_on_off.Text = "0"
        txtZ1b_Blower_on_off.Text = "0"
        txtZ1c1_Blower_on_off.Text = "0"
        txtZ1c2_Blower_on_off.Text = "0"
        txtExhaust_Blower_on_off.Text = "1"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Dryer_On.Click
        ' txtZ1a1_Blower_on_off.Text = InputBox("enter")
        DryerGodet1.Rotation = 270
        AnimationTimer.Enabled = True
        AnimationTimer.Interval = 250

        'Supply_Blower_Z1A.FillColor = Color.Green
        'Z1B_Blower.FillColor = Color.Green
        'Z1C1_Blower.FillColor = Color.Green
        'Z1a2_Blower.FillColor = Color.Green
        'Exhaust_Blower.FillColor = Color.Green
        ' Z1c2_Blower.FillColor = Color.Green
        Humidity_Blower.FillColor = Color.Green

        Air_In_Arrow_1.BlinkMode = 2
        Air_In_Arrow_2.BlinkMode = 2
        Air_In_Arrow_3.BlinkMode = 2
        Air_In_Arrow_4.BlinkMode = 2
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Dryer_Off.Click
        AutomationTimer_Count = 0
        Timer1.Enabled = False
        AnimationTimer.Enabled = False

        Arrow1A.Visible = False
        Arrow1B.Visible = False
        Arrow1C.Visible = False
        Arrow1D.Visible = False
        Arrow2A.Visible = False
        Arrow2B.Visible = False
        Arrow2C.Visible = False
        Arrow2D.Visible = False
        Arrow3A.Visible = False

        Z1a1_Blower.FillColor = Color.Red
        Z1a2_Blower.FillColor = Color.Red
        Z1B_Blower.FillColor = Color.Red
        Z1C1_Blower.FillColor = Color.Red
        Z1c2_Blower.FillColor = Color.Red
        Humidity_Blower.FillColor = Color.Red

        Air_In_Arrow_1.BlinkMode = 0
        Air_In_Arrow_2.BlinkMode = 0
        Air_In_Arrow_3.BlinkMode = 0
        Air_In_Arrow_4.BlinkMode = 0

        txtExhaust_Blower_on_off.Text = "1"

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles AnimationTimer.Tick


        If txtZ1a1_Blower_on_off.Text = "1" Then
            Z1a1_Blower.FillColor = Color.Green
            Z1a1_Blower.Rotation = Z1a1_Blower.Rotation + 90
        Else
            Z1a1_Blower.FillColor = Color.Red
        End If

        If txtZ1a2_Blower_on_off.Text = "1" Then
            Z1a2_Blower.FillColor = Color.Green
            Z1a2_Blower.Rotation = Z1a2_Blower.Rotation + 90
        Else
            Z1a2_Blower.FillColor = Color.Red
        End If

        If txtZ1b_Blower_on_off.Text = "1" Then
            Z1B_Blower.FillColor = Color.Green
            Z1B_Blower.Rotation = Z1B_Blower.Rotation + 90
        Else
            Z1B_Blower.FillColor = Color.Red
        End If

        If txtZ1c1_Blower_on_off.Text = "1" Then
            Z1C1_Blower.FillColor = Color.Green
            Z1C1_Blower.Rotation = Z1C1_Blower.Rotation + 90
        Else
            Z1C1_Blower.FillColor = Color.Red
        End If

        If txtZ1c2_Blower_on_off.Text = "1" Then
            Z1c2_Blower.FillColor = Color.Green
            Z1c2_Blower.Rotation = Z1C1_Blower.Rotation + 90
        Else
            Z1c2_Blower.FillColor = Color.Red
        End If

        If txtExhaust_Blower_on_off.Text = "1" Then
            Exhaust_Blower.FillColor = Color.Green
            Exhaust_Blower.Rotation = Exhaust_Blower.Rotation + 90
        Else
            Exhaust_Blower.FillColor = Color.Red
        End If


        ' Z1B_Blower.Rotation = Z1a1_Blower.Rotation + 90
        ' Z1C1_Blower.Rotation = Z1C1_Blower.Rotation + 90
        ' Exhaust_Blower.Rotation = Exhaust_Blower.Rotation + 90
        ' Z1c2_Blower.Rotation = Z1c2_Blower.Rotation + 90
        Humidity_Blower.Rotation = Humidity_Blower.Rotation + 90


        AutomationTimer_Count = AutomationTimer_Count + 1

        DryerGodet1.Rotation = DryerGodet1.Rotation - 90
        DryerGodet2.Rotation = DryerGodet2.Rotation + 90

        If DryerGodet1.Rotation = 0 Then
            DryerGodet1.Rotation = 270
        End If

        If DryerGodet2.Rotation = 270 Then
            DryerGodet2.Rotation = 0
        End If

        If Z1a1_Blower.Rotation = 270 Then
            Z1a1_Blower.Rotation = 0
        End If

        If Z1a2_Blower.Rotation = 270 Then
            Z1a2_Blower.Rotation = 0
        End If

        If Z1B_Blower.Rotation = 270 Then
            Z1B_Blower.Rotation = 0
        End If

        If Z1C1_Blower.Rotation = 270 Then
            Z1C1_Blower.Rotation = 0
        End If

        If Exhaust_Blower.Rotation = 270 Then
            Exhaust_Blower.Rotation = 0
        End If

        If Z1c2_Blower.Rotation = 270 Then
            Z1c2_Blower.Rotation = 0
        End If

        If Humidity_Blower.Rotation = 270 Then
            Humidity_Blower.Rotation = 0
        End If

        Arrow1A.Visible = True
        Arrow1A.Left = Arrow1A.Left - 10
        Arrow1B.Visible = True
        Arrow1B.Left = Arrow1B.Left - 10
        Arrow1C.Visible = True
        Arrow1C.Left = Arrow1C.Left - 10
        Arrow1D.Visible = True
        Arrow2A.Visible = True
        Arrow2B.Visible = True
        Arrow2B.Left = Arrow2B.Left + 10
        Arrow2C.Visible = True
        Arrow2C.Left = Arrow2C.Left + 10
        Arrow2D.Visible = True
        Arrow3A.Visible = True
        Arrow3A.Left = Arrow3A.Left - 10

        If AutomationTimer_Count >= 4 Then
            Arrow1A.Visible = False
            Arrow1B.Visible = False
            Arrow1C.Visible = False
            Arrow2B.Visible = False
            Arrow2C.Visible = False
            Arrow3A.Visible = False
            Arrow1A.Left = Arrow1A.Left + 40
            Arrow1B.Left = Arrow1B.Left + 40
            Arrow1C.Left = Arrow1C.Left + 40
            Arrow2B.Left = Arrow2B.Left - 40
            Arrow2C.Left = Arrow2C.Left - 40
            Arrow3A.Left = Arrow3A.Left + 40
            AutomationTimer_Count = 0
        End If

        Call CurrentData()
    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Show_Process_Air.Click
        Duct_Cover.Visible = False
        Show_Process_Air.Visible = False
        Hide_Process_Air.Visible = True
    End Sub

    Private Sub Hide_Process_Air_Click(sender As Object, e As EventArgs) Handles Hide_Process_Air.Click
        Duct_Cover.Visible = True
        Hide_Process_Air.Visible = False
        Show_Process_Air.Visible = True
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Show_Hum_Controls.Click
        Access_Door_1.Visible = False
        Access_Door_2.Visible = False
        Access_Door_3.Visible = False
        Access_Door_4.Visible = False
        Access_Door_5.Visible = False
        Access_Door_6.Visible = False

        Show_Hum_Controls.Visible = False
        Hide_Hum_Controls.Visible = True
    End Sub

    Private Sub Hide_Hum_Controls_Click(sender As Object, e As EventArgs) Handles Hide_Hum_Controls.Click
        Access_Door_1.Visible = True
        Access_Door_2.Visible = True
        Access_Door_3.Visible = True
        Access_Door_4.Visible = True
        Access_Door_5.Visible = True
        Access_Door_6.Visible = True

        Show_Hum_Controls.Visible = True
        Hide_Hum_Controls.Visible = False
    End Sub


    Private Sub Animation_Speed_Click(sender As Object, e As EventArgs) Handles Animation_Speed.Click
        AnimationTimer.Interval = InputBox("enter_animation_time")
    End Sub

    Private Sub Recirc_Air_Temp_A_MouseHover(sender As Object, e As EventArgs) Handles Recirc_Air_Temp_A.MouseHover
        txtDevice_1.BringToFront()
        txtDevice_1.Visible = True
        txtDevice_1.Text = "Zone A Temp Sensor"
        txtZ1a_temp_f_pv.BackColor = Color.Cyan
    End Sub


    Private Sub Recirc_Air_Temp_A_MouseLeave(sender As Object, e As EventArgs) Handles Recirc_Air_Temp_A.MouseLeave
        txtDevice_1.SendToBack()
        txtDevice_1.Visible = False
        txtDevice_1.Text = ""
        txtZ1a_temp_f_pv.BackColor = Color.White
    End Sub

    ' Private Sub txtZ1a1_Blower_on_off_TextChanged(sender As Object, e As EventArgs) Handles txtZ1a1_Blower_on_off.TextChanged
    ' If txtZ1a1_Blower_on_off.Text = "1" Then
    '  Z1a1_Blower.FillColor = Color.Green
    ' Else
    '   Z1a1_Blower.FillColor = Color.Red
    ' End If
    ' End Sub

    ' Private Sub txtZ1a2_Blower_on_off_TextChanged(sender As Object, e As EventArgs) Handles txtZ1a2_Blower_on_off.TextChanged
    '  If txtZ1a2_Blower_on_off.Text = "1" Then
    ' Z1a2_Blower.FillColor = Color.Green
    '  Else
    ' Z1a2_Blower.FillColor = Color.Red
    ' End If
    ' End Sub

    Private Sub CurrentData()
        ' current data opens the database and executes the querries
        Dim ConnectionString As String
        Dim connection As SqlConnection
        Dim command As SqlCommand
        Dim Sql As String
        Dim Sql1 As String
        Dim sql2 As String
        Dim Sql3 As String
        Dim Sql4 As String
        Dim sql5 As String
        Dim sql6 As String
        Dim sql7 As String
        Dim sql8 As String
        Dim sql9 As String
        Dim sql10 As String
        Dim sql11 As String
        Dim sql12 As String
        Dim sql13 As String
        Dim sql14 As String
        Dim sql15 As String
        Dim sql16 As String

        ConnectionString = "server = nitta-mfg01\ncimes; Database = Taglogs; User ID = mesuser; Password = nc12m3s0k; Integrated Security = False;"
        Sql = ("select value from[current] where [name] Like 'C3.Dry.inside_casing_wet_flt'")
        Sql1 = ("select value from[current] where [name] like 'C3.Dry.z1a_temp_f_pv'")
        sql2 = ("select value from[current] where [name] like 'C3.Dry.z1b_temp_f_pv'")
        Sql3 = ("select value from[current] where [name] like 'C3.Dry.z1c_temp_f_pv'")
        Sql4 = ("select value from[current] where [name] like 'C3.Dry.z3a_hum_%_pv'")
        sql5 = ("select value from[current] where [name] like 'C3.Dry.z1a1_blw_di'")
        sql6 = ("select value from[current] where [name] like 'C3.Wet.belt_speed_fpm_pv'")
        sql7 = ("select value from[current] where [name] like 'C3.Dry.Line_fpm_pv'")
        sql8 = ("select value from[current] where [name] like 'C3.Dry.casing_temp7_f_pv'")
        sql9 = ("select value from[current] where [name] like 'C3.Dry.z1a2_blw_di'")
        sql10 = ("select value from[current] where [name] like 'C3.Dry.z1b_blw_di'")
        sql11 = ("select value from[current] where [name] like 'C3.Dry.z1c1_blw_di'")
        sql12 = ("select value from[current] where [name] like 'C3.Dry.z1c2_blw_di'")
        sql13 = ("select value from[current] where [name] like 'C3.Dry.casing_temp1_f_pv'")
        sql14 = ("select value from[current] where [name] like 'C3.Dry.casing_temp2_f_pv'")
        sql15 = ("select value from[current] where [name] like 'C3.Dry.casing_temp3_f_pv'")
        sql16 = ("select value from[current] where [name] like 'C3.Dry.casing_temp4_f_pv'")

        connection = New SqlConnection(ConnectionString)
        Try
            connection.Open()
            command = New SqlCommand(Sql, connection)
            Dim SqlReader As SqlDataReader = command.ExecuteReader()

            While SqlReader.Read()
                txtWetness.Text = (SqlReader.Item(0)) ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(Sql1, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1a_temp_f_pv.Text = (SqlReader.Item(0)) & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql2, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1b_temp_f_pv.Text = (SqlReader.Item(0)) & " - " & "Deg F"  ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(Sql3, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1c_temp_f_pv.Text = (SqlReader.Item(0)) & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(Sql4, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ3A_Humidity.Text = (SqlReader.Item(0)) & " % " ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql5, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1a1_Blower_on_off.Text = (SqlReader.Item(0)) '& " % " ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql6, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                Belt_Speed.Text = (SqlReader.Item(0)) & "  FPM " '& SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql7, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                Dryer_speed.Text = (SqlReader.Item(0)) / 10 & "  FPM " ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql8, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtCasing_Temp_7.Text = (SqlReader.Item(0)) / 10 & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql9, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1a2_Blower_on_off.Text = (SqlReader.Item(0))  ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql10, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1b_Blower_on_off.Text = (SqlReader.Item(0))  ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql11, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1c1_Blower_on_off.Text = (SqlReader.Item(0))  ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql12, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtZ1c2_Blower_on_off.Text = (SqlReader.Item(0))  ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql13, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtCasing_temp1.Text = (SqlReader.Item(0)) / 10 & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql14, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtCasing_temp2.Text = (SqlReader.Item(0)) / 10 & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql15, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtCasing_temp3.Text = (SqlReader.Item(0)) / 10 & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command = New SqlCommand(sql16, connection)
            SqlReader = command.ExecuteReader()
            While SqlReader.Read()
                txtCasing_temp4.Text = (SqlReader.Item(0)) / 10 & " - " & "Deg F" ' & "-" & SqlReader.Item(1) & "-" & SqlReader.Item(2))
            End While
            SqlReader.Close()

            command.Dispose()
            connection.Close()

        Catch ex As Exception
            MsgBox("Lost connection to database" + ex.Message)

        End Try
        Return
    End Sub


End Class
