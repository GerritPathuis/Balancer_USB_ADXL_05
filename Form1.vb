Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Text
Imports System.IO
Imports System.IO.Ports
Imports System.ComponentModel
Imports System.Threading
Imports System.Drawing.Graphics
Imports System.Math
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Windows.Forms
Imports System.Drawing.Printing

Public Class Form1
    Dim y_old_s1 As Integer                 'Used in the picturebox4..
    Dim y_old_s2 As Integer                 'Used in the picturebox4
    Dim hertz As Double = 1                 'Hertz of the rotor 
    Dim cntpic3 As Integer                  'Used in the graph
    Dim cntpic4 As Integer                  'Used in the graph
    Dim cntpic5 As Integer                  'Used in the graph
    Dim bulls_eye_counter As Integer = 0
    Dim hoek1, hoek2 As Double              'angle sensor 1 and 2
    Dim weight_left As Double = 1           'weight on left support
    Dim weight_right As Double = 1          'weight on right support
    Dim colorcount As Integer
    Dim sensor_range_ADXL As Double = 12     '[m/s2], Range For the ADXL213 ==> +/- 1.2 g, 
    Dim sensor_range_AS5510 As Double = 2 / 1000.0F             '[m], magnet length is 2.0 [mm]
    Dim bulls_eye_range As Double = 1                           '[mm/2] initial value to prevent zero devision
    Dim AS5510_signal_1, AS5510_signal_2 As Double             ' Between 0.0 and 0.5
    Dim signal_1_strength, signal_2_strength As Double          'signal in 0.0-1.0
    Dim signal_1_strength_mms, signal_2_strength_mms As Double  'signal in [mm/s]
    Dim signal_1_strength_ms2, signal_2_strength_ms2 As Double  'signal on [m/s2]
    Dim calbrate_zero_d1 As Boolean = False
    Dim calbrate_zero_d2 As Boolean = False
    Dim zero_bias_d1 As Double = 0.0    'zero bias sensor1 
    Dim zero_bias_d2 As Double = 0.0    'zero bias sensor2 
    Dim lever_left As Double = 0.34     'accelaration not measured at the center of gravity 
    Dim lever_right As Double = 0.34    'accelaration not measured at the center of gravity 
    Dim g_factor As Double              'G_factor ISO1940 klasse

    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Dim comOpen As Boolean

    Dim m_Bitmap, p_Bitmap, r_Bitmap As Bitmap
    Dim m_Graphics, p_Graphics, q_Graphics As Graphics
    Dim page_no As Integer = 1          'Initial value
    Dim start As Boolean = False

    Private Property ConnectionOK As Boolean
    Private tabControl As TabControl
    Delegate Sub SetTextCallback(ByVal intext As String) 'Added to prevent threading errors during receiveing of data

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False    'Is required to function
        cntpic3 = cntpic4 = cntpic5 = 0
        RadioButton5.Checked = True                    ' set unbalance G6.3
        MakeNewBitmap3()                                'Clear the graph
        Read_settings()                                 'Numeric_Updown_settings
    End Sub

    Private Sub ReceivedText(ByVal intext As String)
        Dim accel1_pro, accel2_pro As Double    '[%]
        Dim y_new1, y_new2, foto_cnt As Integer
        Dim y_new_d1m, y_new_d2m As Integer
        Dim count3 As Integer   'counter for string array
        Dim strArr() As String  'string array
        Dim time_step As Integer = 1 + NumericUpDown3.Value
        Dim path As String = "c:\temp\_mbed.txt" ' Root directory will not work!
        Dim pen3 As New System.Drawing.Pen(Color.Black, 2)
        Dim pen3a As New System.Drawing.Pen(Color.Red, 2)
        Dim abcde(5), tim, ymat As Double
        Dim last_time_measurement As Double
        Dim rad_speed As Double

        Try
            If Me.TxtMbedMessage.InvokeRequired Then
                Dim x As New SetTextCallback(AddressOf ReceivedText)
                Me.Invoke(x, New Object() {(intext)})
            Else
                Me.TxtMbedMessage.AppendText(intext)
            End If
        Catch ex As Exception
            MsgBox("Error 771 received text exception received" & ex.Message)
        End Try

        '-----------------------intext to file -----------------
        If CheckBox1.Checked Then
            If File.Exists(path) = False Then
                Dim createText As String = "Hello and Welcome" & Environment.NewLine    ' Create a file to write to. 
                File.WriteAllText(path, createText)
            End If

            ' This text is always added, making the file longer over time 
            ' if it is not deleted. 
            Dim appendText As String = intext & Environment.NewLine
            File.AppendAllText(path, appendText)
        End If

        '------------------------------------------------
        If ((cntpic3 > PictureBox3.Width - 1) Or (cntpic4 > PictureBox3.Width - 1)) Then
            Drwg_signal_graph()         'Visual check
            cntpic3 = 0                 'Reset counter
            cntpic4 = 0                 'Reset counter
        End If
        '------------------- Sensor #1, picturebox3 actual value for Visual check-----------------------------

        If intext.Contains("D1") And Freeze.Checked = False And Tabs.SelectedTab.Name = "TabPage3" Then
            strArr = intext.Split(" ")
            For count3 = 1 To 1
                accel1_pro = Val(strArr(count3))              'acceleration sensor #1 (0.0 to 1.0 value)
                ' MsgBox("count3= " + count3.ToString + " accel1_pro= " + accel1_pro.ToString)
                accel1_pro = accel1_pro + zero_bias_d1              'verwerking van de bias
                '    MsgBox(accel1_pro)
                '---------------- draw in the graph -----------------------------------------
                If strArr(count3) <> "" And strArr(count3) <> "D1" Then
                    y_new1 = (0.5 - (accel1_pro - 0.5) * Gain.Value * 1.238) * PictureBox3.Height
                    Try
                        p_Graphics.DrawLine(pen3, cntpic3, y_old_s1, cntpic3 + time_step, y_new1)
                    Catch ex As Exception
                        MsgBox("Error 101 writing problem" & ex.Message)
                    End Try
                    cntpic3 = cntpic3 + time_step
                    y_old_s1 = y_new1
                End If
            Next
            PictureBox3.Refresh()
        End If
        '--------------------- Sensor #2, also picturebox3 actual value for visual check------------------------
        If intext.Contains("D2") And Freeze.Checked = False And Tabs.SelectedTab.Name = "TabPage3" Then
            strArr = intext.Split(" ")
            For count3 = 1 To strArr.Length - 1
                accel2_pro = Val(strArr(count3))                'acceleration sensor #2 (-100%/+100%)
                accel2_pro = accel2_pro + zero_bias_d2          'verwerking van de bias

                '---------------- draw in the graph -----------------------------------------
                If strArr(count3) <> "" And strArr(count3) <> "D2" Then
                    y_new2 = (0.5 - (accel2_pro - 0.5) * Gain.Value * 1.238) * PictureBox3.Height
                    Try
                        p_Graphics.DrawLine(pen3a, cntpic4, y_old_s2, cntpic4 + time_step, y_new2)
                    Catch ex As Exception
                        MsgBox("Error 102 writing problem" & ex.Message)
                    End Try
                    cntpic4 = cntpic4 + time_step
                    y_old_s2 = y_new2
                End If
            Next
            PictureBox3.Refresh()
        End If

        '------------------------ calculate the graph and print the resulsts Sensor1 ---------------------
        If intext.Contains("D1M") And Freeze.Checked = False And Tabs.SelectedTab.Name = "TabPage3" Then
            strArr = intext.Split(" ")

            For i = 1 To 5
                abcde(i) = Val(strArr(i))                    'Get a,b,c,d,e,
            Next
            last_time_measurement = Val(strArr(6))           'Get time of last measurement

            'Vertical line for start signal
            p_Graphics.DrawLine(pen3, cntpic3, 0, cntpic3, PictureBox3.Height)
            For i = 0 To 50
                tim = i * last_time_measurement / 50

                ymat = abcde(1) + (abcde(2) * tim) + (abcde(3) * tim ^ 2) + (abcde(4) * tim ^ 3) + (abcde(5) * tim ^ 4)

                '---------------- draw in the graph -----------------------------------------
                y_new_d1m = ymat * Gain.Value * PictureBox3.Height() * 1.238
                Try
                    p_Graphics.DrawLine(pen3, cntpic3, y_old_s1, cntpic3 + time_step, y_new_d1m)
                    cntpic3 = cntpic3 + time_step
                    y_old_s1 = y_new_d1m
                Catch ex As Exception
                    MsgBox("Error 101 writing problem" & ex.Message)
                End Try
            Next
            PictureBox3.Refresh()
        End If

        '------------------------ calculate the graph and print the resulsts Sensor2 ---------------------
        If intext.Contains("D2M") And Freeze.Checked = False And Tabs.SelectedTab.Name = "TabPage3" Then
            strArr = intext.Split(" ")

            For i = 1 To 5
                abcde(i) = Val(strArr(i))
            Next
            last_time_measurement = Val(strArr(6))

            'Vertical line for start signal
            p_Graphics.DrawLine(pen3, cntpic4, 0, cntpic4, PictureBox3.Height)
            For i = 0 To 50
                tim = i * last_time_measurement / 50
                ymat = abcde(1) + (abcde(2) * tim) + (abcde(3) * tim ^ 2) + (abcde(4) * tim ^ 3) + (abcde(5) * tim ^ 4)

                '---------------- draw in the graph -----------------------------------------
                y_new_d2m = ymat * Gain.Value * PictureBox3.Height * 1.238
                Try
                    p_Graphics.DrawLine(pen3a, cntpic4, y_old_s2, cntpic4 + time_step, y_new_d2m)
                    cntpic4 = cntpic4 + time_step
                    y_old_s2 = y_new_d2m
                Catch ex As Exception
                    MsgBox("Error 101 writing problem" & ex.Message)
                End Try
            Next
            PictureBox3.Refresh()
        End If


        '------------------ Results from AS5510 ------------------
        If intext.Contains("AS5510 ") And (Tabs.SelectedTab.Name = "TabPage1" Or Tabs.SelectedTab.Name = "TabPage0") Then
            strArr = intext.Split(" ")
            '---- RPM -------
            hertz = CDbl(Val(strArr(1)))                'Hertz of the rotor
            If hertz <= 0 Then                          'preventing problems
                hertz = 0.1
            End If
            rad_speed = hertz * 2 * PI                  'Radial speed (omega)

            '------------------------- Signals -----------------------------------------
            '!!!!!!!!!!!!!!!!!!This signal is (Top-Top)/2  !!!!!!!!!!!!!!!!!!!!!!!
            '---------------------------------------------------------------------------
            AS5510_signal_1 = Abs(Val(strArr(2)))
            signal_1_strength = AS5510_signal_1 * sensor_range_AS5510             ' radius in [m]
            signal_1_strength_ms2 = signal_1_strength * rad_speed * rad_speed


            AS5510_signal_2 = Abs(Val(strArr(4)))
            signal_2_strength = AS5510_signal_2 * sensor_range_AS5510             ' radius in [m]
            signal_2_strength_ms2 = signal_2_strength * rad_speed * rad_speed

            'MsgBox("AS5510_signal_1 = " & AS5510_signal_1 & " signal_1_strength = " & signal_1_strength & " the signal_1_strength_ms2 = " & signal_2_strength_ms2, vbInformation)

            '--------- angles -------
            hoek1 = Val(strArr(3))                      'v2_max_angle
            hoek2 = Val(strArr(5))                      'v3_max_angle

            '--------- take the sensor positin/location on machine into account -------
            '------------------ compensate +270 for the graph--------------------------
            hoek1 += Sensor_pos.Value + 270
            hoek2 += Sensor_pos.Value + 270
            '----------------------Making sure angle <= 360 degree --------------------
            If hoek1 > 360 Then
                hoek1 -= 360
            End If
            If hoek2 > 360 Then
                hoek2 -= 360
            End If

            'Making sure angle < 0 degree ----------------
            If hoek1 > 360 Then
                hoek1 += 360
            End If
            If hoek2 > 360 Then
                hoek2 += 360
            End If
            '---------------NOW presenting on Bulls Eye--------------------------------
            '--------------------------------------------------------------------------
            bulls_eye_counter = bulls_eye_counter + 1
            If (bulls_eye_counter > 20) Then
                Drwg_bullseye(hertz * 2 * PI)
                bulls_eye_counter = 0
            End If
            If (Tabs.SelectedTab.Name = "TabPage1") Then
                Drwg_bullseye_result(1, signal_1_strength_ms2, hoek1)  'Now draw the result into the left graph
                Drwg_bullseye_result(2, signal_2_strength_ms2, hoek2)  'Now draw the result into the right graph
            End If
        End If

        '-------------------- Results from ADXL ----------
        '----------------------------------------------------------
        If intext.Contains("Results ") And (Tabs.SelectedTab.Name = "TabPage1" Or Tabs.SelectedTab.Name = "TabPage0") Then
            strArr = intext.Split(" ")
            '---- RPM -------
            hertz = CDbl(Val(strArr(1)))                'Hertz of the rotor
            If hertz <= 0 Then                           'preventing problems
                hertz = 0.1
            End If

            '------------------------- Signals -----------------------------------------
            '!!!!!!!!!!!!!!!!!!This signal is (Top-Top)/2  !!!!!!!!!!!!!!!!!!!!!!!
            '---------------------------------------------------------------------------

            signal_1_strength = Abs(Val(strArr(2)))
            signal_1_strength_ms2 = signal_1_strength * sensor_range_ADXL   '[m/s2]

            signal_2_strength = Abs(Val(strArr(4)))
            signal_2_strength_ms2 = signal_2_strength * sensor_range_ADXL   '[m/s2]

            ' MsgBox("signal_1_strength = " & signal_1_strength & " the signal_2_strength = " & signal_2_strength, vbInformation)

            '--------- angles -------
            hoek1 = Val(strArr(3))                      'v2_max_angle
            hoek2 = Val(strArr(5))                      'v3_max_angle

            '--------- take the sensor positin/location on machine into account -------
            '------------------ compensate +270 for the graph--------------------------
            hoek1 += Sensor_pos.Value + 270
            hoek2 += Sensor_pos.Value + 270
            '----------------------Making sure angle <= 360 degree --------------------
            If hoek1 > 360 Then
                hoek1 -= 360
            End If
            If hoek2 > 360 Then
                hoek2 -= 360
            End If

            'Making sure angle < 0 degree ----------------
            If hoek1 > 360 Then
                hoek1 += 360
            End If
            If hoek2 > 360 Then
                hoek2 += 360
            End If
            '---------------NOW presenting on Bulls Eye--------------------------------
            '--------------------------------------------------------------------------
            bulls_eye_counter = bulls_eye_counter + 1
            If (bulls_eye_counter > 20) Then
                Drwg_bullseye(hertz * 2 * PI)
                bulls_eye_counter = 0
            End If
            If (Tabs.SelectedTab.Name = "TabPage1") Then
                Drwg_bullseye_result(1, signal_1_strength_ms2, hoek1)  'Now draw the result into the left graph
                Drwg_bullseye_result(2, signal_2_strength_ms2, hoek2)  'Now draw the result into the right graph
            End If
        End If

        '----------------------------------------------
        If intext.Contains("Photocell") And Tabs.SelectedTab.Name = "TabPage3" Then
            If (cntpic3 > cntpic4) Then     'We want the vertical photocell line at the
                foto_cnt = cntpic3          'highest count is case a sensor is not connected 
            Else
                foto_cnt = cntpic4
            End If
            Try
                p_Graphics.DrawLine(pen3, foto_cnt, 0, foto_cnt, PictureBox3.Height)
            Catch ex As Exception
                MsgBox("Error 103 writing problem" & ex.Message)
            End Try
        End If

    End Sub
    'Clear the bulls eyes
    Private Sub BtnGraphClear_Click(sender As System.Object, e As System.EventArgs) Handles BtnGraphClear.Click
        Drwg_bullseye(hertz * 2 * PI)
    End Sub
    Private Sub Form1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Drwg_bullseye(hertz * 2 * PI)
        Drwg_signal_graph() 'Visual check
    End Sub
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        drwg_balancer_machine()
    End Sub
    'This is the drawing of the Balancing machine
    Private Sub drwg_balancer_machine()
        Dim breed As Integer
        Dim hoog As Integer
        Dim y1, y2, y3, y4 As Integer
        Dim x1, x2 As Integer
        Dim w1, w2 As Integer
        Dim z1 As Integer
        Dim x_sup1, x_sup2 As Integer
        Dim rad1, rad2 As Integer
        Dim clx As Integer
        Dim x_scale, y_scale As Integer

        Dim pen2 As New System.Drawing.Pen(Color.Yellow, 4) 'Used in the graph

        MakeNewBitmap2()    'Clear the graph

        hoog = PictureBox2.Height
        breed = PictureBox2.Width
        x_scale = 9
        y_scale = 10

        Try
            q_Graphics.Clear(Color.Black)  'Clear graph
            'Draw ground
            y1 = hoog - 20
            q_Graphics.DrawLine(pen2, 10, y1, (breed - 10), y1)

            'Draw supports
            y1 = hoog - 20
            y2 = hoog / 2 - 5
            x_sup1 = 40
            q_Graphics.DrawLine(pen2, x_sup1, y1, x_sup1, y2)

            x_sup2 = (rotor_length.Value + 400) / x_scale + x_sup1
            q_Graphics.DrawLine(pen2, x_sup2, y1, x_sup2, y2)

            'Draw center line
            clx = hoog / 2 - 10
            q_Graphics.DrawLine(pen2, x_sup1 - 8, clx, x_sup2 + 8, clx)

            'Draw Rotor radius 1
            rad1 = Radius_left.Value / y_scale
            y1 = clx + rad1
            y2 = clx - rad1
            x1 = cor_pos_left.Value / x_scale + 40
            q_Graphics.DrawLine(pen2, x1, y1, x1, y2)

            'Draw Rotor radius 2
            rad2 = Radius_right.Value / y_scale
            y3 = clx + rad2
            y4 = clx - rad2
            x2 = (400 + rotor_length.Value - cor_pos_rght.Value) / x_scale + 40
            q_Graphics.DrawLine(pen2, x2, y3, x2, y4)

            'Connect the dots
            q_Graphics.DrawLine(pen2, x1, y1, x2, y3)
            q_Graphics.DrawLine(pen2, x1, y2, x2, y4)

            'Add or remove weight
            Dim drawFont As New Font("Arial", 16)
            Dim drawBrush As New SolidBrush(Color.White)
            w1 = x1 - 15
            w2 = x2 - 15

            'Center of gravity
            z1 = x1 + (Center_g.Value * 6) / x_scale

            Dim drawPoint1 As New PointF(w1, 10)    'Correctie gewicht
            Dim drawPoint2 As New PointF(w2, 10)    'Correctie gewicht
            Dim drawPoint3 As New PointF(z1, 30)    'Zwaartepunt

            If RadioButton6.Checked Then
                q_Graphics.DrawString("M+", drawFont, drawBrush, drawPoint1)
            Else
                q_Graphics.DrawString("M-", drawFont, drawBrush, drawPoint1)
            End If
            'Toevoegen-verwijderen gewicht
            If RadioButton9.Checked Then
                q_Graphics.DrawString("M+", drawFont, drawBrush, drawPoint2)
            Else
                q_Graphics.DrawString("M-", drawFont, drawBrush, drawPoint2)
            End If

            'Positie zwaartepunt
            q_Graphics.DrawString("Z", drawFont, drawBrush, drawPoint3)
        Catch ex As Exception
            MsgBox("Error 105 writing problem" & ex.Message)
        End Try
    End Sub

    Private Sub NumericUpDown8_ValueChanged(sender As System.Object, e As System.EventArgs) Handles Radius_left.ValueChanged
        drwg_balancer_machine()
    End Sub
    Private Sub NumericUpDown10_ValueChanged(sender As System.Object, e As System.EventArgs) Handles Radius_right.ValueChanged
        drwg_balancer_machine()
    End Sub
    Private Sub NumericUpDown4_ValueChanged(sender As System.Object, e As System.EventArgs) Handles rotor_length.ValueChanged
        update_rotor_data_screen()
        drwg_balancer_machine()
    End Sub
    Private Sub TabPage3_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles TabPage3.Paint, SplitContainer1.Panel1.Paint
        Drwg_signal_graph() 'Visual check
    End Sub
    Private Sub btnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        Me.SerialPort1.Close()                     'Close existing 
        If cmbPort.Text.Length = 0 Then
            MsgBox("Sorry, did not find any connected USB Balancers")
        Else
            Me.SerialPort1.PortName = cmbPort.Text         'Set SerialPort1 to the selected COM port at startup
            Me.SerialPort1.BaudRate = cmbBaud.Text         'Set Baud rate to the selected value on

            'Other Serial Port Property
            Me.SerialPort1.Parity = IO.Ports.Parity.None
            Me.SerialPort1.StopBits = IO.Ports.StopBits.One
            Me.SerialPort1.DataBits = 8                  'Open our serial port

            Try
                Me.SerialPort1.Open()
                btnConnect.Enabled = False              'Disable Connect button
                btnConnect.BackColor = Color.Yellow
                cmbPort.BackColor = Color.Yellow
                btnConnect.Text = "OK connected"
                btnDisconnect.Enabled = True            'and Enable Disconnect button
            Catch ex As Exception
                MsgBox("Error 654 Open: " & ex.Message)
            End Try

            Try
                Me.SerialPort1.DiscardInBuffer()        'empty inbuffer
                Me.SerialPort1.DiscardOutBuffer()       'empty outbuffer

                Me.SerialPort1.WriteLine("s")           'Real time samples to PC
            Catch ex As Exception
                MsgBox("Error 786 Open: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnDisconnect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisconnect.Click
        Try
            Me.SerialPort1.DiscardInBuffer()
            Me.SerialPort1.Close()             'Close our Serial Port
            Me.SerialPort1.Dispose()
            btnConnect.Enabled = True
            btnConnect.BackColor = Color.Red
            btnConnect.Text = "Connect"
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error 104 Open: " & ex.Message)
        End Try

    End Sub

    Private Sub cmbBaud_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbBaud.SelectedIndexChanged
        If Me.SerialPort1.IsOpen = False Then
            Me.SerialPort1.BaudRate = cmbBaud.Text          'pop a message box to user if he is changing baud rate
        Else                                                'without disconnecting first.
            MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub btnSend_Click(sender As System.Object, e As System.EventArgs) Handles btnSend.Click
        Try
            If Me.SerialPort1.IsOpen = False Then
                Me.SerialPort1.WriteLine(rtbTransmit.Text) 'The text contained in the txtText will be sent to the serial port as ascii
            End If
        Catch exc As IOException
            Console.WriteLine("Error nr 887 IO exception" & exc.Message)
        End Try
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton6.CheckedChanged
        drwg_balancer_machine()
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton9.CheckedChanged
        drwg_balancer_machine()
    End Sub

    Private Sub NumericUpDown9_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cor_pos_rght.ValueChanged
        drwg_balancer_machine()
    End Sub

    Private Sub NumericUpDown5_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cor_pos_left.ValueChanged
        drwg_balancer_machine()
    End Sub

    Private Sub SerialPort1_DataReceived(sender As System.Object, e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try
            Me.ReceivedText(SerialPort1.ReadLine())    'Automatically called every time a data is received at the serialPortb
        Catch exc As IOException
            Console.WriteLine("Error 453 IO exception" & exc.Message)
        End Try
    End Sub

    'setup tab
    Private Sub TabPage0_Enter(sender As System.Object, e As System.EventArgs) Handles TabPage0.Enter
        If (Me.SerialPort1.IsOpen = True) Then ' Preventing exceptions
            'Me.SerialPort1.WriteLine("s")  'Real time samples to PC
            'Me.SerialPort1.DiscardInBuffer()
            TxtMbedMessage.Clear()
        End If
    End Sub
    'Bulls eye tab
    Private Sub TabPage1_Enter_1(sender As Object, e As EventArgs) Handles TabPage1.Enter
        If (Me.SerialPort1.IsOpen = True) Then ' Preventing exceptions
            Me.SerialPort1.WriteLine("m")  'Measure results to PC
            Me.SerialPort1.DiscardInBuffer()
            TxtMbedMessage.Clear()
        End If
    End Sub
    ' graph mbed tab
    Private Sub TabPage3_Enter(sender As System.Object, e As System.EventArgs) Handles TabPage3.Enter
        If (Me.SerialPort1.IsOpen = True) Then ' Preventing exceptions
            Me.SerialPort1.WriteLine("s")  'Real time samples to PC
            Me.SerialPort1.DiscardInBuffer()
        End If
    End Sub
    Private Sub Drwg_bullseye(ByVal omegaa As Double)
        Dim mx1, mx2, mx4, my1, my2, my3, my4, rr, delta_r, le As Integer
        Dim pen1 As New System.Drawing.Pen(Color.Black, 2)
        Dim drawFont As New Font("Arial", 9)
        Dim drawBrush As New SolidBrush(Color.Black)
        Dim strg As Double

        MakeNewBitmap1()    'Clear the graph
        ' Draw Bulls eye #1 -----------------------
        mx1 = PictureBox1.Width * 0.25              'Middelpunt
        mx2 = PictureBox1.Width * 0.75              'Middelpunt
        my1 = PictureBox1.Height * 0.5              'Middelpunt
        my2 = GroupBox7.Height * 0.3                'Label position
        my3 = TabPage1.Height - 20                  'Label position

        mx4 = PictureBox1.Width * 0.5               'Hoek Label position
        my4 = TabPage1.Height - 70                  'Hoek Label position

        Label4.Location = New Point(30, my4)        'Hoek label
        Label6.Location = New Point(mx4 + 30, my4)  'Hoek label

        Label13.Location = New Point(10, my2)
        Label26.Location = New Point(10, my2 + 15)
        Label14.Location = New Point(8, my3)

        If omegaa < 1 Then omegaa = 1 'preventing zero division

        bulls_eye_range = sensor_range_ADXL * 1000 / omegaa / NumericUpDown2.Value   '[mm/s]
        delta_r = PictureBox1.Height / 20           'Radius circles  19->14
        For i = 1 To 7                              'only 7 of the 10 circles are drawn 
            rr = i * delta_r
            Try
                m_Graphics.DrawEllipse(pen1, mx1 - rr, my1 - rr, 2 * rr, 2 * rr)
                m_Graphics.DrawEllipse(pen1, mx2 - rr, my1 - rr, 2 * rr, 2 * rr)

                ' Draw scale
                Dim drawPoint1 As New PointF(mx1, my1 - rr - 12)            'Scale indication
                Dim drawPoint2 As New PointF(mx2, my1 - rr - 12)            'Scale indication
                strg = sensor_range_ADXL * i / (10 * NumericUpDown2.Value)       ' [m/s2]
                strg = strg * 1000 / omegaa                                ' [m/s2-->mm/s]

                If (strg < 1) Then
                    m_Graphics.DrawString(strg.ToString("F2") & " mm/s ", drawFont, drawBrush, drawPoint1)
                    m_Graphics.DrawString(strg.ToString("F2") & " mm/s ", drawFont, drawBrush, drawPoint2)
                Else
                    m_Graphics.DrawString(strg.ToString("F1") & " mm/s ", drawFont, drawBrush, drawPoint1)
                    m_Graphics.DrawString(strg.ToString("F1") & " mm/s ", drawFont, drawBrush, drawPoint2)
                End If

            Catch ex As Exception
                MsgBox("Writing problem 106")
            End Try
        Next i

        ' Draw middle cross
        Try
            le = PictureBox1.Height * 0.4
            m_Graphics.DrawLine(pen1, mx1 - le + 20, my1, mx1 + le - 18, my1) 'Horizontal left
            m_Graphics.DrawLine(pen1, mx1, my1 - le, mx1, my1 + le) 'Vertical left
            m_Graphics.DrawLine(pen1, mx2 - le + 20, my1, mx2 + le - 18, my1) 'Horizontal
            m_Graphics.DrawLine(pen1, mx2, my1 - le, mx2, my1 + le) 'Vertical 

            ' Degree indication left
            m_Graphics.DrawString("0  ==>", drawFont, drawBrush, mx1 - 5, my1 - le - 15)
            m_Graphics.DrawString("90", drawFont, drawBrush, mx1 + le - 18, my1 - 6)
            m_Graphics.DrawString("180", drawFont, drawBrush, mx1 - 13, my1 + le)
            m_Graphics.DrawString("270", drawFont, drawBrush, mx1 - le - 5, my1 - 6)

            ' Degree indication right
            m_Graphics.DrawString("0 ==>", drawFont, drawBrush, mx2 - 5, my1 - le - 15)
            m_Graphics.DrawString("90", drawFont, drawBrush, mx2 + le - 18, my1 - 6)
            m_Graphics.DrawString("180", drawFont, drawBrush, mx2 - 13, my1 + le)
            m_Graphics.DrawString("270", drawFont, drawBrush, mx2 - le - 5, my1 - 6)

        Catch ex As Exception
            MsgBox("Writing problem 107")
        End Try
        PictureBox1.Refresh()
    End Sub

    ' Draw the measured result into the Bulls eye 
    ' accel is given in [mm/s]
    ' omega is given in [rad/sec]
    ' accel_angle is given on [degree, 0-360]
    Private Sub Drwg_bullseye_result(ByVal bullno As Integer, ByVal accel As Double, ByVal accel_angle As Double)
        Dim mx, my, len As Integer
        Dim cx, cy As Integer   'Center x, y
        Dim c_force, c_wght, rpm, velo, omega As Double
        Dim e_radius As Double      'Due to unbalance, shaft rotation with a radius of e_radius [mm] 
        Dim graph_scale As Double
        Dim alfa As Double
        Dim dx, dy As Integer
        Dim pen2r As New System.Drawing.Pen(Color.Red, 5) 'Used in the graph
        Dim pen2g As New System.Drawing.Pen(Color.Green, 4) 'Used in the graph
        Dim drawFont As New Font("Arial", 9)
        Dim drawBrush As New SolidBrush(Color.Black)
        Dim drawBrush2 As New SolidBrush(Color.Black)
        Dim c_veer_L As Double = 200   'N/mm veerconstante opstelling
        Dim c_veer_R As Double = 200   'N/mm veerconstante opstelling
        Dim str, str2 As String

        c_veer_L = Spring_c_L.Value
        c_veer_R = Spring_c_R.Value
        omega = hertz * 2 * PI
        If omega <= 0.1 Then   'Prevent devision by zero
            omega = 0.1
        End If
        '----------- calc-------------
        'Force centri= (massa_rotor/2) x versnelling ()
        'Force_correct= massa_correctie * straal * omega^2

        rpm = hertz * 60.0                  '[rpm]
        velo = 1000 * accel / omega         '[m/s2]-->[mm/s]
        e_radius = velo / omega             'e_radius [mm] (RMS) ===========================

        ' MsgBox("The accel is " & accel & " snelheid is " & velo & " radius is " & e_radius, vbInformation)

        If bullno = 1 Then                                                          'left hand support
            mx = PictureBox1.Width * 0.25                                           'Middelpunt bulls eye #1
            c_force = weight_left * accel                                           'F= m*a  [N]
            c_force = c_force + c_veer_L * e_radius                                 'F= e_radius x veerconstante opstelling
            c_wght = 1000 * 1000 * c_force / (omega * omega * Radius_left.Value)    'Radius in [mm], Correctie gewicht in gram

            str = "L_Sensor= " & accel.ToString("F2") & " [m/s2], v=" & velo.ToString("F2") & " [mm/s] r= " & e_radius.ToString("F4") &
               " [mm], F= " & c_force.ToString("F1") & " [N], C.wght= " & c_wght.ToString("F1") & " [gr], n= " & rpm.ToString("F0") & " [rpm]"
            setlabel13(str)
            str2 = "Hoek= " & accel_angle.ToString("F0")
            setlabel4(str2)
        Else                                                                        'Right hand support   
            mx = PictureBox1.Width * 0.75                                           'Middelpunt Bulls eye #2
            c_force = weight_right * accel                                          'F= m*a [N]
            c_force = c_force + c_veer_R * e_radius                                 'F= e_radius x veerconstante opstelling
            c_wght = 1000 * 1000 * c_force / (omega * omega * Radius_right.Value)   'Radius in [mm], Correctie gewicht in gram

            str = "R_Sensor= " & accel.ToString("F2") & " [m/s2], v=" & velo.ToString("F2") & " [mm/s] r= " & e_radius.ToString("F4") &
                " [mm], F= " & c_force.ToString("F1") & " [N], C.wght= " & c_wght.ToString("F1") & " [gr], n= " & rpm.ToString("F0") & " [rpm]"
            setlabel26(str)
            str2 = "Hoek= " & accel_angle.ToString("F0")
            setlabel6(str2)
        End If

        '--------------------Draw the dot -----------------------------
        my = PictureBox1.Height * 0.5      'Middelpunt
        alfa = accel_angle                 'angle between 0 and 360 degrees

        graph_scale = (velo / bulls_eye_range) * my

        ' MsgBox("zz The velo = " & velo & " the bulls_eye_range = " & bulls_eye_range, vbInformation)


        '---------------- Auto Range --------------------------
        '------------------  down ----------------------------
        'If (velo > bulls_eye_range * 0.8) And (NumericUpDown2.Value > NumericUpDown2.Minimum) Then
        'NumericUpDown2.Value = NumericUpDown2.Value - 2
        'End If
        '------------------- up ----------------------------
        'If (velo < bulls_eye_range * 0.2) And (NumericUpDown2.Value < NumericUpDown2.Maximum - 2) Then
        'NumericUpDown2.Value = NumericUpDown2.Value + 2
        'End If

        dx = Sin((alfa) / 180.0 * PI) * graph_scale
        dy = -Cos((alfa) / 180.0 * PI) * graph_scale
        Try
            If (velo >= g_factor) Then
                len = 6
                cx = mx + dx
                cy = my + dy
                m_Graphics.DrawLine(pen2r, cx - len, cy - len, cx + len, cy + len)      'Cross line 1        
                m_Graphics.DrawLine(pen2r, cx + len, cy - len, cx - len, cy + len)      'Cross line 2
            Else
                m_Graphics.DrawEllipse(pen2g, mx + dx - 3, my + dy - 3, 6, 6)
            End If

        Catch ex As Exception
            MsgBox("Writing problem 108")
        End Try
        PictureBox1.Refresh()
    End Sub
    Private Sub setlabel4(ByVal ptext As String)
        If Me.Label4.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel4)
            Me.Invoke(d, New Object() {(ptext)})
        Else
            Me.Label4.Text = ptext
        End If
    End Sub
    Private Sub setlabel6(ByVal ptext As String)
        If Me.Label6.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel6)
            Me.Invoke(d, New Object() {(ptext)})
        Else
            Me.Label6.Text = ptext
        End If
    End Sub
    Private Sub setlabel7(ByVal ptext As String)
        If Me.Label7.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel7)
            Me.Invoke(d, New Object() {(ptext)})
        Else
            Me.Label7.Text = ptext
        End If
    End Sub
    Private Sub setlabel8(ByVal ptext As String)
        If Me.Label8.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel8)
            Me.Invoke(d, New Object() {(ptext)})
        Else
            Me.Label8.Text = ptext
        End If
    End Sub
    Private Sub setlabel12(ByVal qtext As String)
        If Me.Label12.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel12)
            Me.Invoke(d, New Object() {(qtext)})
        Else
            Me.Label12.Text = qtext
        End If
    End Sub
    Private Sub setlabel13(ByVal qtext As String)
        If Me.Label13.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel13)
            Me.Invoke(d, New Object() {(qtext)})
        Else
            Me.Label13.Text = qtext
        End If
    End Sub
    Private Sub setlabel26(ByVal qtext As String)
        If Me.Label26.InvokeRequired = True Then
            Dim d As New SetTextCallback(AddressOf setlabel26)
            Me.Invoke(d, New Object() {(qtext)})
        Else
            Me.Label26.Text = qtext
        End If
    End Sub
    ' Close the serial port when closing the program
    Private Sub Form1_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        SerialPort1.Close()
        SerialPort1.Dispose()
    End Sub
    ' measure results
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If (RadioButton4.Checked = True) Then
            TxtMbedMessage.Clear()
            Me.SerialPort1.WriteLine("m")       'Measurements to PC
        End If
    End Sub
    'Real time samples
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If (RadioButton3.Checked = True) Then
            TxtMbedMessage.Clear()
            Me.SerialPort1.WriteLine("s")      'Real time samples to PC
        End If
    End Sub
    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles astap_links.ValueChanged
        drwg_balancer_machine()
        update_rotor_data_screen()
    End Sub
    Private Sub NumericUpDown5_ValueChanged_1(sender As Object, e As EventArgs) Handles astap_rechts.ValueChanged
        drwg_balancer_machine()
        update_rotor_data_screen()
    End Sub
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        update_rotor_data_screen()
    End Sub
    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        update_rotor_data_screen()
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        update_rotor_data_screen()
    End Sub
    Private Sub NumericUpDown11_ValueChanged(sender As Object, e As EventArgs) Handles rotor_wght.ValueChanged
        update_rotor_data_screen()
        Calc_resonance()
    End Sub

    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        update_rotor_data_screen()
    End Sub

    Private Sub RadioButton12_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton12.CheckedChanged
        update_rotor_data_screen()
    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged
        update_rotor_data_screen()
    End Sub

    Private Sub update_rotor_data_screen()

        Dim e_per As Double             'Specific residual permissable unbalance
        Dim u_per As Double             'Unbalance permisseble
        Dim centri As Double            'centrifugal force
        Dim omega_max As Double         'hoeksnelheid

        If RadioButton1.Checked Then
            g_factor = 40
        ElseIf RadioButton2.Checked Then
            g_factor = 16
        ElseIf RadioButton5.Checked Then
            g_factor = 6.3
        ElseIf RadioButton7.Checked Then
            g_factor = 2.5
        ElseIf RadioButton11.Checked Then
            g_factor = 1.0
        ElseIf RadioButton12.Checked Then
            g_factor = 0.4
        End If

        Try
            'zwaartepunt werkstuk

            omega_max = 2 * 3.14 * NumericUpDown4.Value / 60
            'e_per = g_factor / hoeksnelheid
            e_per = g_factor * 1000 / omega_max
            Me.setlabel12("E permissable " & e_per.ToString("F1") & " [gr.mm/kg]")

            u_per = e_per * rotor_wght.Value / 2                       'geeft gram.mm
            Me.setlabel7("U permissable " & u_per.ToString("F0") & " [gr.mm] each side")

            centri = u_per / 1000 / 1000 * omega_max * omega_max 'from [mm] and [gr] ---> [m] and [kg]
            Me.setlabel8("Force permissable " & centri.ToString("F0") & " [N] each side")

            'Weight distribution
            weight_left = Center_g.Value / 100 * rotor_wght.Value
            weight_right = rotor_wght.Value - weight_left

            'Versnelling wordt verstrekt door hefboomwerking
            lever_left = (Center_g.Value / 4) / (Center_g.Value / 2 + astap_links.Value)
            lever_right = (Center_g.Value / 4) / (Center_g.Value / 2 + astap_rechts.Value)
        Catch e As DivideByZeroException
        End Try

        Label27.Text = "Gewicht " & weight_left.ToString("F1") & " [kg]"
        Label29.Text = "Gewicht " & weight_right.ToString("F1") & " [kg]"
    End Sub
    Private Sub NumericUpDown12_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        Drwg_bullseye(hertz * 2 * PI)
    End Sub
    ' The sensor give 0-100% equal to -12 to +12 [m/s2]
    ' For visual check 
    ' For setting the bias 
    ' Actual values are added later
    Private Sub Drwg_signal_graph()
        Dim wgr, y1, y2, delta_y As Integer
        Dim pen1 As New System.Drawing.Pen(Color.Black, 1)
        Dim drawFont As New Font("Arial", 9)
        Dim drawBrush As New SolidBrush(Color.Black)
        Dim strg As Double
        Dim hor_lines As Integer

        MakeNewBitmap3()                                        'Clear the graph
        wgr = Me.PictureBox3.Height * 0.5                       'Height

        hor_lines = 10                                          'number horizontal lines in the gragh
        delta_y = PictureBox3.Height / (2 * hor_lines + 2)
        For i = 0 To hor_lines
            y1 = wgr + i * delta_y
            y2 = wgr - i * delta_y
            Try
                p_Graphics.DrawLine(pen1, 75, y1, PictureBox3.Width, y1)
                p_Graphics.DrawLine(pen1, 75, y2, PictureBox3.Width, y2)
            Catch ex As Exception
                MsgBox("Writing problem 109")
            End Try
            ' Draw scale
            Dim drawPoint1 As New PointF(5, y1 - 6)            'Scale indication
            Dim drawPoint2 As New PointF(5, y2 - 6)            'Scale indication
            strg = -1 * sensor_range_ADXL * i / (hor_lines * Gain.Value * 2)
            Try
                p_Graphics.DrawString(strg.ToString("F2") & " [m/s2]", drawFont, drawBrush, drawPoint1)
                strg = strg * -1
                p_Graphics.DrawString(strg.ToString("F2") & " [m/s2]", drawFont, drawBrush, drawPoint2)
            Catch ex As Exception
                MsgBox("Writing problem 110")
            End Try
        Next i
        PictureBox3.Refresh()
    End Sub
    Private Sub gain_ValueChanged(sender As Object, e As EventArgs) Handles Gain.ValueChanged
        Drwg_signal_graph()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        calbrate_zero_d1 = True
        calbrate_zero_d2 = True
    End Sub
    Private Sub PrintPageHandler(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim kantlijn As Integer = 50
        Dim HeaderFont As New Font("Arial", 20)
        Dim date1 As Date = Date.Now
        Dim xx, yy, xxx, yyy As Integer

        Tabs.SelectedTab = TabPage1 'Otherwise it will not show on the print
        Tabs.SelectedTab = TabPage2 'Otherwise it will not show on the print
        Tabs.SelectedTab = TabPage3 'Otherwise it will not show on the print

        Select Case (page_no)
            Case 1                          'Print first page
                PageSetupDialog1.PageSettings.Landscape = False
                e.Graphics.DrawString("BALANCING RESULT SHEET by Balance XS", HeaderFont, Brushes.Black, New Point(kantlijn - 2, 40))
                e.Graphics.DrawString(date1, Me.Font, Brushes.Black, New Point(kantlijn, 70))

                Dim TabBitmap6 As New Bitmap(TabPage1.Width, TabPage1.Height)
                TabPage1.DrawToBitmap(TabBitmap6, New Rectangle(Point.Empty, TabBitmap6.Size))
                e.Graphics.DrawImage(TabBitmap6, kantlijn, 100, 720, 450)

                'e.Graphics.DrawString("Regel3", Me.Font, Brushes.Black, New Point(kantlijn, 545))
                'e.Graphics.DrawString("Regel4", Me.Font, Brushes.Black, New Point(kantlijn, 555))

                TabPage2.DrawToBitmap(TabBitmap6, New Rectangle(Point.Empty, TabBitmap6.Size))
                e.Graphics.DrawImage(TabBitmap6, kantlijn, 600, 720, 450)
                e.HasMorePages = True
                page_no = 2
                TabBitmap6.Dispose()
                Exit Select
            Case 2                          'Print second page
                Dim TabBitmap2 As New Bitmap(TabPage3.Width, TabPage3.Height)
                TabPage3.DrawToBitmap(TabBitmap2, New Rectangle(Point.Empty, TabBitmap2.Size))
                e.Graphics.RotateTransform(90.0F)
                xx = 100    'top margin
                yy = -750   'left margin
                xxx = 900   'Hoogte doc
                yyy = 1400  'Breedte doc
                e.Graphics.DrawImage(TabBitmap2, xx, yy, xx + xxx, yy + yyy)
                e.HasMorePages = False
                page_no = 1
                TabBitmap2.Dispose()
                Exit Select
        End Select
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        PrintDialog1.ShowDialog()
        If PrintPreviewDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
        End If
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        PrintDialog1.ShowDialog()
        If PrintPreviewDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
        End If
    End Sub
    Private Sub MakeNewBitmap1() 'Picture Bull's eye
        Dim wid As Integer = PictureBox1.ClientSize.Width
        Dim hgt As Integer = PictureBox1.ClientSize.Height
        m_Bitmap = New Bitmap(wid, hgt)
        m_Graphics = Graphics.FromImage(m_Bitmap)
        m_Graphics.Clear(Me.BackColor)
        PictureBox1.Image = m_Bitmap
    End Sub
    Private Sub MakeNewBitmap2() 'Balancer picture
        Dim wid As Integer = PictureBox2.ClientSize.Width
        Dim hgt As Integer = PictureBox2.ClientSize.Height
        r_Bitmap = New Bitmap(wid, hgt)
        q_Graphics = Graphics.FromImage(r_Bitmap)
        q_Graphics.Clear(Me.BackColor)
        PictureBox2.Image = r_Bitmap
    End Sub
    Private Sub MakeNewBitmap3() 'Picture signal
        Dim wid As Integer = PictureBox3.ClientSize.Width
        Dim hgt As Integer = PictureBox3.ClientSize.Height
        p_Bitmap = New Bitmap(wid, hgt)
        p_Graphics = Graphics.FromImage(p_Bitmap)
        p_Graphics.Clear(Me.BackColor)
        PictureBox3.Image = p_Bitmap
    End Sub
    Private Sub Serial_setup() 'Serial port setup
        If (Me.SerialPort1.IsOpen = True) Then ' Preventing exceptions
            Me.SerialPort1.DiscardInBuffer()
            Me.SerialPort1.Close()
        End If

        Try
            Me.myPort = SerialPort.GetPortNames() 'Get all com ports available
            For Each port In myPort
                Me.cmbPort.Items.Add(port)
            Next port
            Me.cmbPort.Text = cmbPort.Items.Item(0)    'Set cmbPort text to the first COM port detected
        Catch ex As Exception
            MsgBox("No com ports detected")
        End Try

        Me.cmbBaud.Items.Add(9600)     'Populate the cmbBaud Combo box to common baud rates used
        Me.cmbBaud.Items.Add(19200)
        Me.cmbBaud.Items.Add(38400)
        Me.cmbBaud.Items.Add(57600)
        Me.cmbBaud.Items.Add(115200)
        Me.cmbBaud.Items.Add(230400)
        Me.cmbBaud.SelectedIndex = 5    'Set cmbBaud text to 230400 Baud 

        Me.SerialPort1.ReceivedBytesThreshold = 24    'wait EOF char or until there are x bytes in the buffer, include \n and \r !!!!
        Me.SerialPort1.ReadBufferSize = 4096
        Me.SerialPort1.DiscardNull = True              'important otherwise it will not work
        Me.SerialPort1.Parity = Parity.None
        Me.SerialPort1.StopBits = StopBits.One
        Me.SerialPort1.Handshake = Handshake.None
        Me.SerialPort1.ParityReplace = True
        btnDisconnect.Enabled = False                  'Initially Disconnect Button is Disabled
    End Sub

    Private Sub cmbPort_Click(sender As Object, e As EventArgs) Handles cmbPort.Click
        cmbPort.SelectedIndex = -1
        cmbPort.Items.Clear()
        Serial_setup()
    End Sub
    Private Sub NumericUpDown5_ValueChanged_2(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged
        zero_bias_d1 = NumericUpDown5.Value / 10000
    End Sub
    Private Sub NumericUpDown6_ValueChanged_1(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged
        zero_bias_d2 = NumericUpDown6.Value / 10000
    End Sub
    Private Sub Store_settings()
        Dim path2 As String = "c:\temp\_mbed_settings.txt" ' Root directory will not work!
        Dim createText As String
        Try
            createText = Gain.Value & Environment.NewLine    ' Create a file to write to. 
            createText = createText & NumericUpDown3.Value & Environment.NewLine
            createText = createText & NumericUpDown5.Value & Environment.NewLine
            createText = createText & NumericUpDown6.Value & Environment.NewLine
            File.WriteAllText(path2, createText)
        Catch e As Exception
            Console.WriteLine("Problem writing file: ")
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Private Sub Read_settings()
        Dim path2 As String = "c:\temp\_mbed_settings.txt" ' Root directory will not work!

        Try
            Using sr As New StreamReader(path2)
                Gain.Value = CDec(sr.ReadLine())
                NumericUpDown3.Value = CDec(sr.ReadLine())
                NumericUpDown5.Value = CDec(sr.ReadLine())
                NumericUpDown6.Value = CDec(sr.ReadLine())
            End Using
        Catch e As Exception
            Console.WriteLine("The file could not be read:")
            Console.WriteLine(e.Message)
        End Try
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Store_settings()
    End Sub
    Private Sub Calc_resonance()
        Dim spr As Double
        Dim wht As Double
        Dim resn As Double
        spr = Spring_c_L.Value
        wht = rotor_wght.Value
        If (spr > 0 And wht > 0) Then
            resn = Sqrt(4 * Spring_c_L.Value / rotor_wght.Value) * 60 / (2 * PI)

            Label2.Text = "Resonance at " + resn.ToString("f0") + " [rpm]"
        End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles Spring_c_L.ValueChanged
        Calc_resonance()
    End Sub
    'Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    '    Drwg_signal_graph()
    ' End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.SerialPort1.WriteLine("q")           'Real time samples to PC from middel connector
        Catch ex As Exception
            MsgBox("Error 991 Open: " & ex.Message)
        End Try
    End Sub
    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        Try
            Me.SerialPort1.WriteLine("t")           'Real time samples to PC from middel connector
        Catch ex As Exception
            MsgBox("Error 99a Open: " & ex.Message)
        End Try
    End Sub

    Private Sub Center_g_ValueChanged(sender As Object, e As EventArgs) Handles Center_g.ValueChanged
        drwg_balancer_machine()
    End Sub

End Class