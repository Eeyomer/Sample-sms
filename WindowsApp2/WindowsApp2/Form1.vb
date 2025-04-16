Imports System.Net.Mail
Imports System.IO.Ports
Imports System.Threading
Imports System.Net

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeSerialPort()
    End Sub ' Create a serial port object
    Dim WithEvents serialPort As New SerialPort()

    ' Initialize SerialPort settings
    Sub InitializeSerialPort()
        Try
            ' Set the COM port name (change to your USB-to-Serial COM port)
            serialPort.PortName = "COM9"  ' Update with your COM port
            serialPort.BaudRate = 9600     ' Baud rate for SIM800C
            serialPort.Parity = Parity.None
            serialPort.StopBits = StopBits.One
            serialPort.DataBits = 8
            serialPort.NewLine = vbCrLf  ' Newline character for message separation
            serialPort.Handshake = Handshake.None
            serialPort.Encoding = System.Text.Encoding.ASCII

            ' Open the serial port
            serialPort.Open()
            lblStatus.Text = "Serial Port Opened Successfully."

        Catch ex As Exception
            lblStatus.Text = "Error opening serial port: " & ex.Message
        End Try
    End Sub

    ' Send AT command to the GSM module
    Sub SendATCommand(command As String)
        Try
            If serialPort.IsOpen Then
                ' Write the AT command to the serial port
                serialPort.WriteLine(command)
                lblStatus.Text = "Sent: " & command
            End If
        Catch ex As Exception
            lblStatus.Text = "Error sending command: " & ex.Message
        End Try
    End Sub

    ' Read response from the GSM module
    Sub ReadResponse()
        Try
            If serialPort.IsOpen Then
                ' Read the response from the GSM module
                Dim response As String = serialPort.ReadLine()
                lblStatus.Text = "Response: " & response
            End If
        Catch ex As Exception
            lblStatus.Text = "Error reading response: " & ex.Message
        End Try
    End Sub

    ' Send SMS using the provided phone number and message
    Sub SendSMS(phoneNumber As String, message As String)
        Try
            If serialPort.IsOpen Then
                ' Set the message format to text mode
                SendATCommand("AT+CMGF=1")
                Thread.Sleep(1000) ' Wait for the command to take effect

                ' Send the phone number and the message
                SendATCommand("AT+CMGS=""" & phoneNumber & """")
                Thread.Sleep(1000) ' Wait for the command to take effect
                serialPort.WriteLine(message)  ' Write the message content
                serialPort.WriteLine(Chr(26))  ' Send Ctrl+Z to indicate end of message
                lblStatus.Text = "Message Sent Successfully!"
            End If
        Catch ex As Exception
            lblStatus.Text = "Error sending SMS: " & ex.Message
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim phoneNumber As String = TextBox1.Text
        Dim message As String = TextBox2.Text

        ' Validate inputs
        If String.IsNullOrEmpty(phoneNumber) Or String.IsNullOrEmpty(message) Then
            lblStatus.Text = "Please enter both phone number and message."
        Else
            ' Send SMS
            SendSMS(phoneNumber, message)
        End If
    End Sub

    Private Sub BtnSendEmail_Click(sender As Object, e As EventArgs) Handles BtnSendEmail.Click
        Try
            ' Hardcoded sender email and app password
            Dim senderEmail As String = "renzirish.garcia@gmail.com"
            Dim appPassword As String = "mnjd dmjd gwow wxpg"

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.EnableSsl = True
            smtp.Credentials = New NetworkCredential(senderEmail, appPassword)

            ' Create the email message
            Dim mail As New MailMessage()
            mail.From = New MailAddress(senderEmail)
            mail.To.Add(TxtReceiver.Text)
            mail.Subject = "KUPSSSS"
            mail.Body = TxtEmail.Text  ' txtEmail holds the content/body

            ' Send the email
            smtp.Send(mail)
            MessageBox.Show("Email sent successfully!")


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
