Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal value As String, ByVal worksection As Integer)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents input_emailbody As System.Windows.Forms.RichTextBox
    Friend WithEvents input_coursecode As System.Windows.Forms.TextBox
    Friend WithEvents input_emailsender As System.Windows.Forms.TextBox
    Friend WithEvents validstudents As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents validcourses As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents totalresults As System.Windows.Forms.Label
    Friend WithEvents emailssent As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents processstart As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents processend As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents input_emailsubject As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label6 = New System.Windows.Forms.Label
        Me.input_emailsubject = New System.Windows.Forms.TextBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.input_emailbody = New System.Windows.Forms.RichTextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.input_emailsender = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.input_coursecode = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label8 = New System.Windows.Forms.Label
        Me.processend = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.processstart = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.emailssent = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.totalresults = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.validcourses = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.validstudents = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Commerce Course Mass Mailer"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Monitor Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.input_emailsubject)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Button3)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.input_emailsender)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.input_coursecode)
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Location = New System.Drawing.Point(8, 8)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(664, 360)
        Me.Panel2.TabIndex = 43
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(192, 16)
        Me.Label6.TabIndex = 54
        Me.Label6.Text = "Email Subject"
        '
        'input_emailsubject
        '
        Me.input_emailsubject.BackColor = System.Drawing.Color.Bisque
        Me.input_emailsubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.input_emailsubject.ForeColor = System.Drawing.Color.Black
        Me.input_emailsubject.Location = New System.Drawing.Point(16, 72)
        Me.input_emailsubject.Name = "input_emailsubject"
        Me.input_emailsubject.Size = New System.Drawing.Size(328, 20)
        Me.input_emailsubject.TabIndex = 3
        Me.input_emailsubject.Text = ""
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.input_emailbody)
        Me.Panel3.Location = New System.Drawing.Point(16, 120)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(328, 224)
        Me.Panel3.TabIndex = 52
        '
        'input_emailbody
        '
        Me.input_emailbody.BackColor = System.Drawing.Color.Bisque
        Me.input_emailbody.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.input_emailbody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.input_emailbody.ForeColor = System.Drawing.Color.Black
        Me.input_emailbody.Location = New System.Drawing.Point(0, 0)
        Me.input_emailbody.Name = "input_emailbody"
        Me.input_emailbody.Size = New System.Drawing.Size(326, 222)
        Me.input_emailbody.TabIndex = 4
        Me.input_emailbody.Text = ""
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.White
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Red
        Me.Button3.Location = New System.Drawing.Point(352, 304)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(40, 40)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "GO!"
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(152, 11)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(192, 16)
        Me.Label14.TabIndex = 50
        Me.Label14.Text = "Email Sender"
        '
        'input_emailsender
        '
        Me.input_emailsender.BackColor = System.Drawing.Color.Bisque
        Me.input_emailsender.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.input_emailsender.ForeColor = System.Drawing.Color.Black
        Me.input_emailsender.Location = New System.Drawing.Point(152, 27)
        Me.input_emailsender.Name = "input_emailsender"
        Me.input_emailsender.Size = New System.Drawing.Size(192, 20)
        Me.input_emailsender.TabIndex = 2
        Me.input_emailsender.Text = ""
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(16, 104)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 16)
        Me.Label13.TabIndex = 48
        Me.Label13.Text = "Email Body"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 46
        Me.Label4.Text = "Course Code"
        '
        'input_coursecode
        '
        Me.input_coursecode.BackColor = System.Drawing.Color.Bisque
        Me.input_coursecode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.input_coursecode.ForeColor = System.Drawing.Color.Black
        Me.input_coursecode.Location = New System.Drawing.Point(16, 27)
        Me.input_coursecode.Name = "input_coursecode"
        Me.input_coursecode.Size = New System.Drawing.Size(120, 20)
        Me.input_coursecode.TabIndex = 1
        Me.input_coursecode.Text = ""
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Bisque
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.processend)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.processstart)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.emailssent)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.totalresults)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.validcourses)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.validstudents)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.PictureBox5)
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(400, 184)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(248, 160)
        Me.Panel1.TabIndex = 43
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Gray
        Me.Label8.Location = New System.Drawing.Point(16, 128)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 44
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'processend
        '
        Me.processend.ForeColor = System.Drawing.Color.Gray
        Me.processend.Location = New System.Drawing.Point(88, 88)
        Me.processend.Name = "processend"
        Me.processend.Size = New System.Drawing.Size(145, 16)
        Me.processend.TabIndex = 43
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Gray
        Me.Label11.Location = New System.Drawing.Point(8, 88)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(112, 16)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Process End:"
        '
        'processstart
        '
        Me.processstart.ForeColor = System.Drawing.Color.Gray
        Me.processstart.Location = New System.Drawing.Point(88, 8)
        Me.processstart.Name = "processstart"
        Me.processstart.Size = New System.Drawing.Size(152, 16)
        Me.processstart.TabIndex = 41
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Gray
        Me.Label9.Location = New System.Drawing.Point(8, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Process Start:"
        '
        'emailssent
        '
        Me.emailssent.ForeColor = System.Drawing.Color.Gray
        Me.emailssent.Location = New System.Drawing.Point(88, 72)
        Me.emailssent.Name = "emailssent"
        Me.emailssent.Size = New System.Drawing.Size(152, 16)
        Me.emailssent.TabIndex = 39
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(8, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 38
        Me.Label7.Text = "Emails Sent:"
        '
        'totalresults
        '
        Me.totalresults.ForeColor = System.Drawing.Color.Gray
        Me.totalresults.Location = New System.Drawing.Point(88, 24)
        Me.totalresults.Name = "totalresults"
        Me.totalresults.Size = New System.Drawing.Size(152, 16)
        Me.totalresults.TabIndex = 37
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Gray
        Me.Label10.Location = New System.Drawing.Point(8, 24)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(112, 16)
        Me.Label10.TabIndex = 36
        Me.Label10.Text = "Total Results:"
        '
        'validcourses
        '
        Me.validcourses.ForeColor = System.Drawing.Color.Gray
        Me.validcourses.Location = New System.Drawing.Point(88, 56)
        Me.validcourses.Name = "validcourses"
        Me.validcourses.Size = New System.Drawing.Size(152, 16)
        Me.validcourses.TabIndex = 35
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Gray
        Me.Label5.Location = New System.Drawing.Point(8, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Valid Courses:"
        '
        'validstudents
        '
        Me.validstudents.ForeColor = System.Drawing.Color.Gray
        Me.validstudents.Location = New System.Drawing.Point(88, 40)
        Me.validstudents.Name = "validstudents"
        Me.validstudents.Size = New System.Drawing.Size(152, 16)
        Me.validstudents.TabIndex = 33
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Gray
        Me.Label3.Location = New System.Drawing.Point(8, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Valid Students:"
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(208, 128)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 26
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(192, 128)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 25
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(176, 128)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 24
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(160, 128)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(144, 128)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(8, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(232, 32)
        Me.Label2.TabIndex = 28
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(368, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(288, 160)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "Commerce Courses Mass Mailer is an application that can be used to send text emai" & _
        "ls to students registered for various Commerce courses. The application works by" & _
        " searching through objects on the Novell Network Tree of the University of Cape " & _
        "Town (UCT) network via LDAP (Light-weight Directory Access Protocol) calls, matc" & _
        "hing Course Groups on the user inputted filter, and extracting a list of student" & _
        " numbers from these located course groups. Once the student number list has been" & _
        " extracted, a list of email addresses is composed and the inputted text email is" & _
        " mailed to each address using UCT's mail server."
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Orange
        Me.ClientSize = New System.Drawing.Size(682, 376)
        Me.Controls.Add(Me.Panel2)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(688, 408)
        Me.MinimumSize = New System.Drawing.Size(688, 408)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "Commerce Courses Mass Mailer"
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False



    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Commerce Courses Mass Mailer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

   


    Private Sub run_green_lights()
        Try
            '  Label1.ForeColor = Color.LimeGreen
            '     Label1.Text = "Monitoring"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while changing status light colour. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            'Label1.ForeColor = Color.DarkOrange
            '  Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while changing status light colour. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the status light strip. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the status light strip. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")

            ' Label11.Text = "0"

            Timer2.Start()


        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the monitoring screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Worker1.Dispose()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing the Monitoring screen. The program will attempt to recover shortly.")
        End Try
    End Sub



    Public Sub WorkerHandler(ByVal Result As String)
        Try
            currently_working = False
            NotifyIcon1.Text = "Commerce Course Mass Mailer"
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal value As String, ByVal worksection As Integer)
        Try
            Select Case worksection
                Case 1
                    processstart.Text = value
                    processstart.Refresh()
                Case 2
                    totalresults.Text = value
                    totalresults.Refresh()
                Case 3
                    validstudents.Text = value
                    validstudents.Refresh()
                Case 4
                    validcourses.Text = value
                    validcourses.Refresh()
                Case 5
                    emailssent.Text = value
                    emailssent.Refresh()
                Case 6
                    processend.Text = value
                    processend.Refresh()
            End Select
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub run_worker()
        Try

        
        If currently_working = False Then
                run_orange_lights()
                Worker1.emailbody = input_emailbody.Text.Trim
                Worker1.emailsender = input_emailsender.Text.Trim
                Worker1.emailsubject = input_emailsubject.Text.Trim
                Worker1.coursecode = input_coursecode.Text.Trim
                Worker1.ChooseThreads(1)
                currently_working = True
            End If
        Catch ex As Exception
            Error_Handler(ex.ToString)
        End Try
    End Sub







    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Worker1.Dispose()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing Running Applications Monitor. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while trying to display the main screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while trying to display the main screen. The program will attempt to recover shortly.")
        End Try
    End Sub







    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (Not input_emailbody.Text = "" And Not input_emailsender.Text = "" And Not input_emailsubject.Text = "" And Not input_coursecode.Text = "") Then
            run_worker()
        End If
    End Sub


End Class
