﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.TxtEmail = New System.Windows.Forms.RichTextBox()
        Me.TxtReceiver = New System.Windows.Forms.TextBox()
        Me.BtnSendEmail = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(267, 70)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(319, 22)
        Me.TextBox1.TabIndex = 0
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(267, 208)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(319, 22)
        Me.TextBox2.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(413, 266)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(228, 363)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(51, 17)
        Me.lblStatus.TabIndex = 3
        Me.lblStatus.Text = "Label1"
        '
        'TxtEmail
        '
        Me.TxtEmail.Location = New System.Drawing.Point(883, 208)
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(416, 96)
        Me.TxtEmail.TabIndex = 4
        Me.TxtEmail.Text = ""
        '
        'TxtReceiver
        '
        Me.TxtReceiver.Location = New System.Drawing.Point(883, 124)
        Me.TxtReceiver.Name = "TxtReceiver"
        Me.TxtReceiver.Size = New System.Drawing.Size(319, 22)
        Me.TxtReceiver.TabIndex = 5
        '
        'BtnSendEmail
        '
        Me.BtnSendEmail.Location = New System.Drawing.Point(1050, 323)
        Me.BtnSendEmail.Name = "BtnSendEmail"
        Me.BtnSendEmail.Size = New System.Drawing.Size(75, 23)
        Me.BtnSendEmail.TabIndex = 6
        Me.BtnSendEmail.Text = "Button2"
        Me.BtnSendEmail.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1395, 450)
        Me.Controls.Add(Me.BtnSendEmail)
        Me.Controls.Add(Me.TxtReceiver)
        Me.Controls.Add(Me.TxtEmail)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents lblStatus As Label
    Friend WithEvents TxtEmail As RichTextBox
    Friend WithEvents TxtReceiver As TextBox
    Friend WithEvents BtnSendEmail As Button
End Class
