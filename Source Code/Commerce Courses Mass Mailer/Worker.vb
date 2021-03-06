Imports System.IO
Imports System.Web.Mail
Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public coursecode As String = ""
    Public emailsubject As String = ""
    Public emailbody As String = ""
    Public emailsender As String = ""

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal value As String, ByVal worksection As Integer)
    'worksection: 1 - process start
    'worksection: 2 - total results
    'worksection: 3 - valid students
    'worksection: 4 - valid courses
    'worksection: 5 - emails sent
    'worksection: 6 - process end

#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Commerce Courses Mass Mailer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Public Sub WorkerExecute()
        Try
            RaiseEvent WorkerProgress("", 1)
            RaiseEvent WorkerProgress("0", 2)
            RaiseEvent WorkerProgress("0", 3)
            RaiseEvent WorkerProgress("0", 4)
            RaiseEvent WorkerProgress("0", 5)
            RaiseEvent WorkerProgress("", 6)

            RaiseEvent WorkerProgress(Format(Now(), "dd/MM/yyyy hh:mm:ss tt").ToString, 1)
            Dim error_encounter As Boolean = False
            Dim error_message As String = ""

            Dim ProcID As Integer
            Dim apppath As String = Application.StartupPath


            If apppath.EndsWith("\") Then
                apppath = apppath.Remove(apppath.Length - 1, 1)
            End If

            If System.IO.File.Exists(apppath & "\result.txt") = True Then
                System.IO.File.Delete(apppath & "\result.txt")
            End If
            Dim runprog As String = """" & apppath & "\Commerce Courses Student List Extractor.exe"" """ & coursecode & """ > """ & apppath & "\result.txt"""
            DosShellCommand(runprog)


            Dim arrstudents As ArrayList = New ArrayList
            Dim arrcourses As ArrayList = New ArrayList

            arrstudents.Clear()
            arrcourses.Clear()

            Dim fin As FileInfo = New FileInfo(apppath & "\result.txt")
            If fin.Exists = True Then
                Dim filereader As StreamReader = New StreamReader(apppath & "\result.txt")
                Dim read As String
                Dim cour As String
                Dim stu As String
                While filereader.Peek > -1
                    read = filereader.ReadLine().Trim
                    If read.StartsWith("Status Code") = False And read.StartsWith("Error Notice") = False And Not read = "" And Not read Is Nothing Then
                        cour = read.Split(" ").GetValue(0)
                        stu = read.Split(" ").GetValue(1)
                        If stu.IndexOf(",") < 0 And IsNumeric(stu.Substring(stu.Length - 2, 1)) = True Then
                            arrstudents.Add(stu)
                            arrcourses.Add(cour)
                        End If
                    End If

                End While
                filereader.Close()
                filereader = Nothing
            End If
            arrcourses.Sort()
            arrstudents.Sort()

            RaiseEvent WorkerProgress(arrstudents.Count.ToString, 2)

            Dim cc As String = ""
            Dim dd As String = ""
            Dim il As Long


            RaiseEvent WorkerProgress(arrstudents.Count.ToString, 3)
            For il = arrstudents.Count - 1 To 0 Step -1
                dd = arrstudents.Item(il)
                If dd = cc Then
                    arrstudents.RemoveAt(il)
                    RaiseEvent WorkerProgress(arrstudents.Count.ToString, 3)
                Else
                    cc = dd
                End If
            Next

            cc = ""
            dd = ""
            RaiseEvent WorkerProgress(arrcourses.Count.ToString, 4)
            For il = arrcourses.Count - 1 To 0 Step -1
                dd = arrcourses.Item(il)
                If dd = cc Then
                    arrcourses.RemoveAt(il)
                    RaiseEvent WorkerProgress(arrcourses.Count.ToString, 4)
                Else
                    cc = dd
                End If
            Next

            Dim result As DialogResult
            result = MsgBox(arrstudents.Count & " unique mail adresses were located. Do you still wish to send out your email to this group?", MsgBoxStyle.OKCancel, "Send Email Confirmation")
            If result = DialogResult.OK Then
                For il = 0 To arrstudents.Count - 1
                    Try
                        If TextMail("mail.uct.ac.za", emailsender, arrstudents.Item(il) & "@mail.uct.ac.za", emailsubject, emailbody) = True Then
                            RaiseEvent WorkerProgress(il + 1, 5)
                        End If
                    Catch ex As Exception
                        Error_Handler("An """ & ex.Message & """ error occurred while sending a notification email. The program will attempt to recover shortly.")
                    End Try
                Next

            Else
                RaiseEvent WorkerProgress(0, 5)
            End If



            result = "Success"
            RaiseEvent WorkerProgress(Format(Now(), "dd/MM/yyyy hh:mm:ss tt").ToString, 6)
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerProgress(Format(Now(), "dd/MM/yyyy hh:mm:ss tt").ToString, 6)
            RaiseEvent WorkerComplete(result)
        End Try


    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching DOS shell. The program will attempt to recover shortly.")
        End Try
        Return s
    End Function

    Public Function TextMail(ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1
                objMail = New MailMessage
                objMail.BodyFormat = MailFormat.Text
                objMail.From = "webserver@commerce.uct.ac.za"
                objMail.To = emailaddys(counter).Trim
                objMail.Subject = strSubj
                objMail.Body = strBody

                SmtpMail.SmtpServer = "mail.uct.ac.za"
                SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

    Public Function TextMail(ByVal SmtpServer As String, ByVal strFrom As String, ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1


                objMail = New MailMessage
                objMail.BodyFormat = MailFormat.Text
                objMail.From = strFrom
                objMail.To = emailaddys(counter).Trim
                objMail.Subject = strSubj
                objMail.Body = strBody

                SmtpMail.SmtpServer = SmtpServer
                SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

End Class
