Public Class Form1
    Inherits System.Windows.Forms.Form
    Private ThirdOctave() As Double = {100, 125, 160, 200, 250, 316, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3160, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000}
    Private UPV As New UPx.Application       'Include UPV driver
    Private Structure Curve                                                 'Stores one curve for writing them to the data base later
        Dim XArray() As Double
        Dim YArray() As Double
    End Structure
    Private Structure ValuePair
        Dim Value1 As Double
        Dim Value2 As Double
    End Structure
    Private Const NaN As Double = 3.40282E+38
    Private Cancel As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown3 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown4 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents ReflectAutoCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents ReverbAutoCheckBox As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.NumericUpDown4 = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.NumericUpDown3 = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button5 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.ReflectAutoCheckBox = New System.Windows.Forms.CheckBox
        Me.ReverbAutoCheckBox = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ReverbAutoCheckBox)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 192)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reverberation Time"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Test Signal"
        '
        'RadioButton3
        '
        Me.RadioButton3.Checked = True
        Me.RadioButton3.Location = New System.Drawing.Point(256, 24)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton3.TabIndex = 10
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Random"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(176, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton2.TabIndex = 9
        Me.RadioButton2.Text = "Multisine"
        '
        'RadioButton1
        '
        Me.RadioButton1.Location = New System.Drawing.Point(96, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton1.TabIndex = 8
        Me.RadioButton1.Text = "Sine"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(168, 160)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Cancel Measurement"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 160)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(120, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Start Measurement"
        '
        'ComboBox2
        '
        Me.ComboBox2.Items.AddRange(New Object() {"100 Hz", "125 Hz", "160 Hz", "200 Hz", "250 Hz", "316 Hz", "400 Hz", "500 Hz", "630 Hz", "800 Hz", "1 kHz", "1.25 kHz", "1.6 kHz", "2 kHz", "2.5 kHz", "3.16 kHz", "4 kHz", "5 kHz", "6.3 kHz", "8 kHz", "10 kHz", "12.5 kHz", "16 kHz", "20 kHz"})
        Me.ComboBox2.Location = New System.Drawing.Point(232, 96)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(88, 21)
        Me.ComboBox2.TabIndex = 5
        Me.ComboBox2.Text = "select ..."
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(208, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "to"
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"100 Hz", "125 Hz", "160 Hz", "200 Hz", "250 Hz", "316 Hz", "400 Hz", "500 Hz", "630 Hz", "800 Hz", "1 kHz", "1.25 kHz", "1.6 kHz", "2 kHz", "2.5 kHz", "3.16 kHz", "4 kHz", "5 kHz", "6.3 kHz", "8 kHz", "10 kHz", "12.5 kHz", "16 kHz", "20 kHz"})
        Me.ComboBox1.Location = New System.Drawing.Point(112, 96)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(88, 21)
        Me.ComboBox1.TabIndex = 3
        Me.ComboBox1.Text = "select ..."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 99)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Third Octaves from"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(176, 64)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {60, 0, 0, -2147483648})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(72, 20)
        Me.NumericUpDown1.TabIndex = 1
        Me.NumericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.NumericUpDown1.Value = New Decimal(New Integer() {60, 0, 0, -2147483648})
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(160, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "RMS Output Voltage (dBV)"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ReflectAutoCheckBox)
        Me.GroupBox2.Controls.Add(Me.NumericUpDown4)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Controls.Add(Me.NumericUpDown3)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 232)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(338, 176)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Reflectogram"
        '
        'NumericUpDown4
        '
        Me.NumericUpDown4.Location = New System.Drawing.Point(184, 88)
        Me.NumericUpDown4.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NumericUpDown4.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown4.Name = "NumericUpDown4"
        Me.NumericUpDown4.Size = New System.Drawing.Size(72, 20)
        Me.NumericUpDown4.TabIndex = 11
        Me.NumericUpDown4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.NumericUpDown4.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(160, 16)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "No. of Averages"
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(168, 144)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(120, 23)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Cancel Measurement"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(40, 144)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(120, 23)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "Start Measurement"
        '
        'NumericUpDown3
        '
        Me.NumericUpDown3.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumericUpDown3.Location = New System.Drawing.Point(184, 55)
        Me.NumericUpDown3.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.NumericUpDown3.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown3.Name = "NumericUpDown3"
        Me.NumericUpDown3.Size = New System.Drawing.Size(72, 20)
        Me.NumericUpDown3.TabIndex = 5
        Me.NumericUpDown3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.NumericUpDown3.Value = New Decimal(New Integer() {500, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(160, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "High Pass Filter @ Hz"
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Location = New System.Drawing.Point(184, 23)
        Me.NumericUpDown2.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown2.Minimum = New Decimal(New Integer() {60, 0, 0, -2147483648})
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(72, 20)
        Me.NumericUpDown2.TabIndex = 3
        Me.NumericUpDown2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.NumericUpDown2.Value = New Decimal(New Integer() {60, 0, 0, -2147483648})
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Peak Output Voltage (dBV)"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(176, 472)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(120, 23)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Close"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(112, 208)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(224, 20)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = "Idle"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 208)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Status"
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(24, 416)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.TabIndex = 10
        Me.CheckBox1.Text = "Live Update"
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(168, 440)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(168, 24)
        Me.CheckBox2.TabIndex = 11
        Me.CheckBox2.Text = "Minimize when Completed"
        '
        'CheckBox3
        '
        Me.CheckBox3.Location = New System.Drawing.Point(168, 416)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(176, 24)
        Me.CheckBox3.TabIndex = 12
        Me.CheckBox3.Text = "Minimize during Measurement"
        '
        'ReflectAutoCheckBox
        '
        Me.ReflectAutoCheckBox.Checked = True
        Me.ReflectAutoCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ReflectAutoCheckBox.Location = New System.Drawing.Point(16, 112)
        Me.ReflectAutoCheckBox.Name = "ReflectAutoCheckBox"
        Me.ReflectAutoCheckBox.TabIndex = 12
        Me.ReflectAutoCheckBox.Text = "Autoscale"
        '
        'ReverbAutoCheckBox
        '
        Me.ReverbAutoCheckBox.Checked = True
        Me.ReverbAutoCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ReverbAutoCheckBox.Location = New System.Drawing.Point(16, 128)
        Me.ReverbAutoCheckBox.Name = "ReverbAutoCheckBox"
        Me.ReverbAutoCheckBox.TabIndex = 13
        Me.ReverbAutoCheckBox.Text = "Autoscale"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 501)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Room Acoustics Measurements"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Function InitiateUpx() As Boolean
        Dim Index As Byte
        Dim Result As Boolean = False
        Dim Counter As Integer = 0
        Dim Instring As String

        Try
            Debug.Write("Trying to initialize UPV remote control...")
            UPV.InitTCP("localhost") ' connect UPV
            UPV.Timeout = 10000
            Do                                                                                    'Read error queue until "0; No Error" shows up
                UPV.Write("SYST:ERR?")
                Instring = UPV.Read
                If Instring.StartsWith("0") Then
                    Debug.WriteLine(" succeeded")
                    Result = True
                End If
                Counter += 1
            Loop Until Result Or Counter > 10000

        Catch ex As Exception
            MessageBox.Show("Initialization of UPV remote control failed." & ControlChars.CrLf & "Check whether UPV firmware is running and try again!")
            Application.Exit()
        End Try
        Return Result
    End Function

    Private Sub CloseUpx()
        UPV.Close() ' disconnect UPV
    End Sub

    Private Sub UPVCommand(ByVal CommandString As String)   'Sends command to UPV and checks error queue
        Dim ErrString As String

        UPV.Write(CommandString)
        Debug.Write("UPV Command = " & CommandString)
        UPV.Write("Syst:Err?")
        Do
            ErrString = UPV.Read
            If ErrString = String.Empty Then
                Debug.Write("Empty error string received: ")
                Debug.WriteLine(UPV.ErrorString)
            End If
        Loop Until ErrString <> String.Empty                'workaround for occasional empty strings
        If Not ErrString.StartsWith("0") Then
            MessageBox.Show("UPV Error" & ControlChars.CrLf & ErrString, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Application.Exit()
        End If
        Debug.WriteLine("   Err = " & ErrString)
    End Sub

    Private Sub UPVCommand(ByVal CommandString As String, ByVal Deb As Boolean)   'Sends command to UPV and checks error queue
        Dim ErrString As String

        UPV.Write(CommandString)
        If Deb Then
            Debug.Write("UPV Command = " & CommandString)
        Else
            Debug.Write("UPV Command sent. Length = " & CommandString.Length.ToString)
        End If
        UPV.Write("Syst:Err?")
        Do
            ErrString = UPV.Read
            If ErrString = String.Empty Then
                Debug.Write("Empty error string received: ")
                Debug.WriteLine(UPV.ErrorString)
            End If
        Loop Until ErrString <> String.Empty                'workaround for occasional empty strings
        If Not ErrString.StartsWith("0") Then
            MessageBox.Show("UPV Error" & ControlChars.CrLf & ErrString, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Application.Exit()
        End If
        Debug.WriteLine("   Err = " & ErrString)
    End Sub

    Private Function UPVQuery(ByVal QueryString As String) As String        'Sends a query to UPV, reads response and checks error queue
        Dim AnswerString As String
        Dim ErrString As String

        Try
            UPV.Write(QueryString)
            Debug.Write("UPV Query = " & QueryString)
            Delay(50)
            Do
                AnswerString = UPV.Read
                If AnswerString = String.Empty Then
                    Debug.WriteLine("   Empty answer string received" & "  UPV.ErrorString = " & UPV.ErrorString)
                    UPV.Write(QueryString)
                    Debug.Write("UPV Query = " & QueryString)
                End If
            Loop Until AnswerString <> String.Empty             'workaround for occasional empty strings
            Debug.Write("   Answer = " & AnswerString)
            UPV.Write("Syst:Err?")
            Do
                ErrString = UPV.Read
                If ErrString = String.Empty Then
                    Debug.WriteLine(" Empty error string received")
                End If
            Loop Until ErrString <> String.Empty
            If Not ErrString.StartsWith("0") Then
                MessageBox.Show("UPV Error" & ControlChars.CrLf & ErrString, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Application.Exit()
            End If
            Debug.WriteLine("   Err = " & ErrString)
            Return AnswerString
        Catch ex As Exception
            MessageBox.Show("UPV query error")
        End Try
    End Function

    Private Function UPVQuery(ByVal QueryString As String, ByVal Deb As Boolean) As String        'Sends a query to UPV, reads response and checks error queue
        Dim AnswerString As String
        Dim ErrString As String

        Try
            UPV.Write(QueryString)
            Debug.Write("UPV Query = " & QueryString)
            Delay(50)
            Do
                AnswerString = UPV.Read
                If AnswerString = String.Empty Then
                    Debug.WriteLine("   Empty answer string received" & "  UPV.ErrorString = " & UPV.ErrorString)
                    UPV.Write(QueryString)
                    Debug.Write("UPV Query = " & QueryString)
                End If
            Loop Until AnswerString <> String.Empty             'workaround for occasional empty strings
            If Deb Then
                Debug.Write("   Answer = " & AnswerString)
            Else
                Debug.Write("   Answer received, length = " & AnswerString.Length.ToString)
            End If
            UPV.Write("Syst:Err?")
            Do
                ErrString = UPV.Read
                If ErrString = String.Empty Then
                    Debug.WriteLine(" Empty error string received")
                End If
            Loop Until ErrString <> String.Empty
            If Not ErrString.StartsWith("0") Then
                MessageBox.Show("UPV Error" & ControlChars.CrLf & ErrString, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Application.Exit()
            End If
            Debug.WriteLine("   Err = " & ErrString)
            Return AnswerString
        Catch ex As Exception
            MessageBox.Show("UPV query error")
        End Try
    End Function

    Private Sub Delay(ByVal DelayTime As Long)
        Dim TimeInstance As Long
        TimeInstance = System.Environment.TickCount
        Do
            Application.DoEvents()
        Loop Until (System.Environment.TickCount - TimeInstance > DelayTime)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 Then
            If ComboBox1.SelectedIndex > ComboBox2.SelectedIndex Then ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox1.SelectedIndex > -1 Then
            If ComboBox2.SelectedIndex < ComboBox1.SelectedIndex Then ComboBox2.SelectedIndex = ComboBox1.SelectedIndex
        End If
    End Sub

    Private Function Average(ByVal Array() As Double) As Double                           'Returns average of all values in the array
        Dim Index As Integer
        Dim Result As Double = 0
        For Index = 0 To Array.Length - 1
            Result = Result + Array(Index)
        Next
        Result = Result / Array.Length
        Return Result
    End Function

    Private Sub Regression(ByVal XArray() As Double, ByVal YArray() As Double, ByRef Slope As Double, ByRef Intercept As Double, ByRef Deviation As Double)      'Calculates straight regression line through points specified by XArray and YArray

        Dim Count As Integer = XArray.Length
        Dim N As Double = 0
        Dim Z As Double = 0
        Dim XAverage As Double = Average(XArray)
        Dim YAverage As Double = Average(YArray)
        Dim S As Double = 0
        Dim Index As Integer

        Try
            For Index = 0 To Count - 1
                Z = Z + (XArray(Index) * YArray(Index))
                N = N + (XArray(Index) * XArray(Index))
            Next
            Z = Z - (Count * XAverage * YAverage)
            N = N - (Count * XAverage * XAverage)
            Slope = Z / N
            Intercept = YAverage - Slope * XAverage
            For Index = 0 To Count - 1
                S = S + Slope * Slope * (XArray(Index) - XAverage) * (XArray(Index) - XAverage) - 2 * Slope * (XArray(Index) - XAverage) * (YArray(Index) - YAverage) + (YArray(Index) - YAverage) * (YArray(Index) - YAverage)
            Next
            Deviation = S / Count
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Function SubArray(ByVal OriginalArray() As Double, ByVal StartIndex As Integer, ByVal Length As Integer) As Double()           'Copies the specified number of element to a new array, starting with the first element
        Dim Result() As Double
        Dim Index As Integer

        If (OriginalArray.Length > 0) And (Length <= OriginalArray.Length) Then
            ReDim Result(Length - 1)
            For Index = StartIndex To StartIndex + Length - 1
                Result(Index - StartIndex) = OriginalArray(Index)
            Next
        End If
        Return Result
    End Function

    Private Function ValField(ByVal valuelist As String) As Double()
        Dim Splitted() As String
        Dim Index1 As Long
        Dim Index2 As Integer = 0
        Dim Field() As Double

        Splitted = Split(valuelist, ",")                            'Split String
        ReDim Field(Splitted.Length - 1)
        For Index1 = LBound(Splitted) To UBound(Splitted)
            Field(Index1) = Val(Splitted(Index1))              'Enter value into field
        Next Index1
        Return Field
    End Function

    Private Function ReverbTime(ByVal ThirdOct As Byte) As ValuePair
        Dim Result As ValuePair
        Dim Level() As Double
        Dim Time() As Double
        Dim Instring As String
        Dim SigEnd As Integer
        Dim SubLength As Integer = 1
        Dim XSubarray() As Double
        Dim YSubarray() As Double
        Dim Slope() As Double
        Dim Intercept() As Double
        Dim Residue() As Double
        Dim Bend As Boolean
        Dim Index As Integer
        Dim StartIndex As Integer
        Dim OriginalLength As Integer
        Dim OnLevel As Double
        Dim Count As Byte
        Dim SigEndTime = 1.25
        Dim GenVolt As Double

        Try
            UPVCommand("SENSe:UFILter6:CENTer " & ThirdOctave(ThirdOct).ToString & " HZ")       'Set generator filter frequency
            If RadioButton1.Checked Then    'Sinewave
                UPVCommand("SOURce:FREQuency " & ThirdOctave(ThirdOct).ToString & " HZ")
            End If
            If RadioButton2.Checked Then    'Multisine
                UPVCommand("SOURce:RANDom:SPACing:FREQuency " & (0.08 * ThirdOctave(ThirdOct)).ToString & " HZ")
                UPVCommand("SENSe:VOLTage:APERture " & (12.5 / ThirdOctave(ThirdOct)).ToString & " S")
                UPVCommand("SOURce:RANDom:FREQuency:LOWer " & (ThirdOctave(ThirdOct) / 1.122).ToString & " HZ")           'Set analyzer frequency
                UPVCommand("SOURce:RANDom:FREQuency:UPPer " & (ThirdOctave(ThirdOct) * 1.122).ToString & " HZ")           'Set analyzer frequency
            End If
            If RadioButton3.Checked Then    'Random
                If ThirdOctave(ThirdOct) < 1000 Then
                    UPVCommand("SENSe:VOLTage:APERture " & (20 / ThirdOctave(ThirdOct)).ToString & " S")
                End If
                GenVolt = NumericUpDown1.Value + 10 * Math.Log10(20000 / (0.23 * ThirdOctave(ThirdOct)))
                If GenVolt > 20 Then
                    MessageBox.Show("Required generator voltage exceeds range." & ControlChars.CrLf & " Reduce RMS Output Voltage and try again!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Cancel = True
                    Exit Function
                Else
                    UPVCommand("SOURce:VOLTage:RMS " & GenVolt.ToString & " DBV") 'Set RMS voltage
                End If
            End If

            Application.DoEvents()
            UPVCommand("OUTPUT ON")                                                   'Start measurement
            UPVCommand("INIT:CONT OFF; *WAI")                                                   'Start measurement
            Instring = UPVQuery("TRAC:SWE1:LOAD:AX?")                                                             'Level over time
            Time = ValField(Instring)                                                          'Convert to array
            Instring = UPVQuery("TRAC:SWE1:LOAD:AY?")                                                             'Level over time
            Level = ValField(Instring)                                                          'Convert to array
            OnLevel = 0
            Application.DoEvents()
            For Index = 0 To Level.Length - 1
                If Time(Index) > 0.5 And Time(Index) < 1 Then
                    OnLevel += Level(Index)
                    Count += 1
                End If
                Application.DoEvents()
            Next
            OnLevel = OnLevel / Count
            Debug.WriteLine("OnLevel = " & OnLevel.ToString)

            '****************** Determine reverb time ******************
            Debug.WriteLine("Determining reverb time")
            OriginalLength = 380
            ReDim Slope(OriginalLength)
            ReDim Intercept(OriginalLength)
            ReDim Residue(OriginalLength)
            If ThirdOct = 0 Then SigEndTime = 1.75
            If (ThirdOct > 0) And (ThirdOct < 5) Then SigEndTime = 1.5
            SigEnd = 0
            Do
                SigEnd += 1
            Loop Until Time(SigEnd) > SigEndTime
            Debug.WriteLine("SigEnd = " & SigEnd.ToString & "  Time = " & Time(SigEnd.ToString))
            SubLength = 1
            Do
                SubLength += 1
                XSubarray = SubArray(Time, SigEnd, SubLength)
                YSubarray = SubArray(Level, SigEnd, SubLength)
                Regression(XSubarray, YSubarray, Slope(SubLength), Intercept(SubLength), Residue(SubLength))           'Calculate straight regression line
                Debug.WriteLine("SubLength = " & SubLength.ToString & "  Slope = " & Slope(SubLength).ToString & "  Residue = " & Residue(SubLength).ToString)
                If SubLength > 5 Then
                    Bend = (Residue(SubLength) > Residue(SubLength - 1))                       'curve fit deteriorates if the curve bends
                    Bend = Bend And (Residue(SubLength - 2) > Residue(SubLength - 3))
                    Bend = Bend And (Residue(SubLength - 3) > Residue(SubLength - 4))
                    Bend = Bend And (Residue(SubLength - 4) > Residue(SubLength - 5))
                    Bend = Bend And (Residue(SubLength - 5) > Residue(SubLength - 6))
                    Bend = Bend And (Slope(SubLength) > Slope(SubLength - 1))               'resulting slope gets flatter
                    Bend = Bend And (Slope(SubLength) > Slope(SubLength - 2))
                    Bend = Bend And (Slope(SubLength) > Slope(SubLength - 3))             'resulting slope gets flatter
                    Bend = Bend And (Slope(SubLength) > Slope(SubLength - 4))
                End If
            Loop Until (Bend And (SubLength > 7)) Or SubLength = Time.Length - SigEnd - 2        'Assure that there are enough points to fit straight regression line
            Index = SubLength - 6
            Result.Value1 = -60 / Slope(Index)                                                    'Settling time
            Debug.WriteLine("Reverb Time = " & Result.Value1.ToString)


        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
        Return Result
    End Function

    Private Function Assemble(ByVal ValueArray() As Double) As String
        Dim Result As String = String.Empty
        Dim Index1 As Integer
        Dim MaxInd As Integer = ValueArray.Length - 1

        Try
            Result = ValueArray(0).ToString
            For Index1 = 1 To MaxInd
                If Index1 < ValueArray.Length Then
                    Result = Result & "," & ValueArray(Index1).ToString("#.#####")
                End If
            Next
            Return Result
        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

    Private Function Assemble(ByVal ValueArray() As Double, ByVal Factor As Double) As String
        Dim Result As String = String.Empty
        Dim Index1 As Integer
        Dim MaxInd As Integer = ValueArray.Length - 1

        Try
            Result = (Math.Abs(ValueArray(0)) * Factor).ToString
            For Index1 = 1 To MaxInd
                If Index1 < ValueArray.Length Then
                    Result = Result & "," & (Math.Abs(ValueArray(Index1)) * Factor).ToString("0.00000E+00")
                End If
            Next
            Return Result
        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

    Private Sub SendCurve(ByVal Trace As Curve, ByVal TraceB As Boolean)
        Dim Outstring As String

        UPVCommand("DISP:SWE2:X:SOUR MAN")
        UPVCommand("DISP:SWE2:X:SCAL MAN")
        UPVCommand("DISPlay:SWEep2:X:LEFT " & Trace.XArray(0).ToString & " S")
        UPVCommand("DISPlay:SWEep2:X:RIGHT " & Trace.XArray(Trace.XArray.Length - 1).ToString & " S")
        If TraceB Then
            Outstring = Assemble(Trace.XArray)
            UPVCommand("TRAC:SWE2:STORE:BX " & Outstring)
            Outstring = Assemble(Trace.YArray)
            UPVCommand("TRAC:SWE2:STORE:BY " & Outstring)
        Else
            Outstring = Assemble(Trace.XArray)
            UPVCommand("TRAC:SWE2:STORE:AX " & Outstring)
            Outstring = Assemble(Trace.YArray)
            UPVCommand("TRAC:SWE2:STORE:AY " & Outstring)
        End If

    End Sub

    Private Sub SendWav(ByVal Trace() As Double, ByVal Factor As Double, ByVal TraceB As Boolean)
        Dim Outstring As String

        If TraceB Then
            Outstring = Assemble(Trace, Factor)
            UPVCommand("TRAC:WAV1:STORE:BY " & Outstring, False)
        Else
            Outstring = Assemble(Trace, Factor)
            UPVCommand("TRAC:WAV1:STORE:AY " & Outstring, True)
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click     'Measure reverberation time
        Dim Index1 As Byte
        Dim RevTime As Curve
        Dim SettlingTime As Curve
        Dim RT As ValuePair
        Dim CurveMax As Double = 0

        Try
            Button1.Enabled = False
            Button2.Enabled = True
            Button4.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            RadioButton3.Enabled = False

            TextBox1.Text = "Configuring UPV"
            Delay(10)
            UPVCommand("MMEMory:LOAD:STATe 'C:/Program Files/Rohde&Schwarz/1GA61/Reverb.set'")  'Load Setup
            If RadioButton1.Checked Then    'Sine
                UPVCommand("SOURce:FUNCtion BURSt")
                UPVCommand("SOURce:VOLTage " & NumericUpDown1.Value.ToString & " DBV")      'Set RMS voltage
            End If
            If RadioButton2.Checked Then    'Multisine
                UPVCommand("SOURce:FUNCtion RANDom")
                UPVCommand("SOURce:RANDom:DOMain FREQuency")
                UPVCommand("SOURce:FILTer OFF")
                UPVCommand("SOURce:VOLTage:RMS " & NumericUpDown1.Value.ToString & " DBV")      'Set RMS voltage
            End If
            If RadioButton3.Checked Then    'Random
                UPVCommand("SOURce:FUNCtion RANDom")
                UPVCommand("SOURce:RANDom:DOMain TIME")
                UPVCommand("SOURce:FILTer UFIL6")
                UPVCommand("SOURce:VOLTage:RMS " & NumericUpDown1.Value.ToString & " DBV")      'Set RMS voltage
            End If

            ReDim RevTime.XArray(ComboBox2.SelectedIndex - ComboBox1.SelectedIndex)         'Number of 3rd octaves
            ReDim RevTime.YArray(ComboBox2.SelectedIndex - ComboBox1.SelectedIndex)
            ReDim SettlingTime.XArray(ComboBox2.SelectedIndex - ComboBox1.SelectedIndex)
            ReDim SettlingTime.YArray(ComboBox2.SelectedIndex - ComboBox1.SelectedIndex)

            UPVCommand("OUTPUT OFF")

            If CheckBox1.Checked Then
                UPVCommand("SYST:DISP:SCP ON")
            Else
                UPVCommand("SYST:DISP:SCP OFF")
            End If

            If CheckBox3.Checked Then
                Me.WindowState = FormWindowState.Minimized
            End If

            For Index1 = ComboBox1.SelectedIndex To ComboBox2.SelectedIndex                 'Measure reverb times and enter them to result curves
                TextBox1.Text = "Measuring reverb time @ " & ThirdOctave(Index1).ToString & "Hz"
                Delay(10)
                Application.DoEvents()
                RT = ReverbTime(Index1)
                Application.DoEvents()
                RevTime.XArray(Index1 - ComboBox1.SelectedIndex) = ThirdOctave(Index1)
                RevTime.YArray(Index1 - ComboBox1.SelectedIndex) = RT.Value1
                If RT.Value1 > CurveMax Then CurveMax = RT.Value1
                Application.DoEvents()
                If Cancel Then Exit For
            Next

            UPVCommand("OUTPUT OFF")

            If Not Cancel Then
                UPVCommand("MMEMory:LOAD:STATe 'C:/Program Files/Rohde&Schwarz/1GA61/Reverb_Result.set'")  'Load Setup
                SendCurve(RevTime, False)                                                    'Enter results into Sweep Graph 2
                If ReverbAutoCheckBox.Checked Then UPVCommand("DISPlay:SWEep2:A:TOP " & (1.05 * CurveMax).ToString & "  V")
                If CheckBox2.Checked Then
                    Me.WindowState = FormWindowState.Minimized
                Else
                    Me.WindowState = FormWindowState.Normal
                End If
            Else
                Cancel = False
            End If


            Button1.Enabled = True                                                          'Enable buttons
            Button2.Enabled = False
            Button4.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True

            UPV.Write("*GTL")                                                               'Send UPV to local
            TextBox1.Text = "Idle"
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            MessageBox.Show("An error occurred during reverberation time measurement", "Measurement Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitiateUpx()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        CloseUpx()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Cancel = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Average As Integer
        Dim Instring As String
        Dim AverageCurve() As Double
        Dim Splitted() As String
        Dim Sample As Integer
        Dim CurveMax As Double = 0

        Try
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = True
            Button4.Enabled = False

            TextBox1.Text = "Configuring UPV"
            Delay(10)

            UPVCommand("MMEMory:LOAD:STATe 'C:/Program Files/Rohde&Schwarz/1GA61/Reflectogram.set'")
            UPVCommand("SENSe:UFILter2:PASSb " & NumericUpDown3.Value.ToString & " HZ")
            UPVCommand("SOURce:VOLTage:TOTal " & NumericUpDown2.Value.ToString & " DBV")
            UPVCommand("SENSe7:TRIGger:LEVel " & (NumericUpDown2.Value - 20).ToString & " DBV")
            UPVCommand("SENSe:VOLTage:RANGe2:VALue " & (Math.Pow(10, NumericUpDown2.Value / 20)).ToString)
            UPVCommand("DISPlay:WAVeform:A:BOTTom -1 mV")

            UPVCommand("OUTPUT ON")

            If CheckBox1.Checked Then
                UPVCommand("SYST:DISP:SCP ON")
            Else
                UPVCommand("SYST:DISP:SCP OFF")
            End If

            If CheckBox3.Checked Then
                Me.WindowState = FormWindowState.Minimized
            End If

            For Average = 1 To NumericUpDown4.Value
                TextBox1.Text = "Average No. " & Average.ToString
                Delay(10)
                UPVCommand("INIT:CONT OFF;*WAI")
                Instring = UPVQuery("TRAC:WAV:LOAD:AY?", False)         'Load waveform
                Splitted = Instring.Split(","c)
                If Average = 1 Then
                    ReDim AverageCurve(Splitted.Length - 1)             'initialize array for average
                    For Sample = 0 To Splitted.Length - 1
                        AverageCurve(Sample) = 0
                    Next
                End If
                For Sample = 0 To Splitted.Length - 1
                    AverageCurve(Sample) += Val(Splitted(Sample))       'add to average
                    If AverageCurve(Sample) > CurveMax Then CurveMax = AverageCurve(Sample)
                Next
                If Cancel Then Exit For
            Next

            UPVCommand("OUTPUT OFF")

            If Not Cancel Then
                TextBox1.Text = "Scaling output"
                Delay(10)
                SendWav(AverageCurve, 1 / NumericUpDown4.Value, False)
                UPVCommand("DISPlay:WAVeform:A:BOTTom 0 V")
                If ReflectAutoCheckBox.Checked Then UPVCommand("DISPlay:WAVeform:A:TOP " & (1.05 * CurveMax / NumericUpDown4.Value).ToString & "  V")
                If CheckBox2.Checked Then
                    Me.WindowState = FormWindowState.Minimized
                Else
                    Me.WindowState = FormWindowState.Normal
                End If
            Else
                Cancel = False
            End If

            UPV.Write("*GTL")
            TextBox1.Text = "Idle"

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Cancel = True
    End Sub
End Class
