Imports System.Security.Principal
Imports System.Security.Permissions
Imports System.Runtime.InteropServices
Imports System.Environment
Imports System.IO

Public Class Login_Screen
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents lblErrorMessage As System.Windows.Forms.Label
    Friend WithEvents btnLogin As System.Windows.Forms.Button
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Login_Screen))
        Me.lblErrorMessage = New System.Windows.Forms.Label
        Me.btnLogin = New System.Windows.Forms.Button
        Me.txtUsername = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblErrorMessage
        '
        Me.lblErrorMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErrorMessage.ForeColor = System.Drawing.Color.Gold
        Me.lblErrorMessage.Location = New System.Drawing.Point(224, 72)
        Me.lblErrorMessage.Name = "lblErrorMessage"
        Me.lblErrorMessage.Size = New System.Drawing.Size(192, 56)
        Me.lblErrorMessage.TabIndex = 0
        '
        'btnLogin
        '
        Me.btnLogin.BackColor = System.Drawing.Color.MidnightBlue
        Me.btnLogin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogin.Location = New System.Drawing.Point(24, 112)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.TabIndex = 1
        Me.btnLogin.Text = "LOGIN"
        '
        'txtUsername
        '
        Me.txtUsername.BackColor = System.Drawing.Color.White
        Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUsername.ForeColor = System.Drawing.Color.Black
        Me.txtUsername.Location = New System.Drawing.Point(24, 40)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(184, 18)
        Me.txtUsername.TabIndex = 2
        Me.txtUsername.Text = ""
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.Color.White
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPassword.ForeColor = System.Drawing.Color.Black
        Me.txtPassword.Location = New System.Drawing.Point(24, 80)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(184, 18)
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Username:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblErrorMessage)
        Me.GroupBox1.Controls.Add(Me.btnLogin)
        Me.GroupBox1.Controls.Add(Me.txtUsername)
        Me.GroupBox1.Controls.Add(Me.txtPassword)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(424, 152)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(224, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(192, 32)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "TO LOGIN TO THE SYSTEM, PLEASE ENTER YOUR WINDOWS USER ACCOUNT  DETAILS AND HIT T" & _
        "HE 'LOGIN' BUTTON TO PROCEED."
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Password:"
        '
        'Login_Screen
        '
        Me.AcceptButton = Me.btnLogin
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 11)
        Me.BackColor = System.Drawing.Color.LightSlateGray
        Me.ClientSize = New System.Drawing.Size(440, 174)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.White
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(448, 208)
        Me.MinimumSize = New System.Drawing.Size(448, 208)
        Me.Name = "Login_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Windows Authentication Required"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    'The LogonUser function tries to log on to the local computer
    'by using the specified user name. The function authenticates
    'the Windows user with the password provided.
    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
       ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
       ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
       ByRef phToken As IntPtr) As Boolean

    'The FormatMessage function formats a message string that is passed as input.
    <DllImport("kernel32.dll")> _
    Public Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, _
       ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], _
       ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function

    'The CloseHandle function closes the handle to an open object such as an Access token.
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean

    'The GetErrorMessage function formats and then returns an error message
    'that corresponds to the input error code.
    Public Shared Function GetErrorMessage(ByVal errorCode As Integer) As String
        Dim FORMAT_MESSAGE_ALLOCATE_BUFFER As Integer = &H100
        Dim FORMAT_MESSAGE_IGNORE_INSERTS As Integer = &H200
        Dim FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000

        Dim msgSize As Integer = 255
        Dim lpMsgBuf As String
        Dim dwFlags As Integer = FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS

        Dim lpSource As IntPtr = IntPtr.Zero
        Dim lpArguments As IntPtr = IntPtr.Zero
        'Call the FormatMessage function to format the message.
        Dim returnVal As Integer = FormatMessage(dwFlags, lpSource, errorCode, 0, lpMsgBuf, _
                msgSize, lpArguments)
        If returnVal = 0 Then
            Throw New Exception("Failed to format message for error code " + errorCode.ToString() + ". ")
        End If
        Return lpMsgBuf
    End Function

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim tokenHandle As New IntPtr(0)
        Try
            If txtUsername.Text.Trim = "" Then
                lblErrorMessage.Text = "Login failure: You need to enter a valid username"
                Exit Try
            End If
            Dim UserName, MachineName, Pwd As String
            'The MachineName property gets the name of your computer.
            MachineName = System.Environment.MachineName
            UserName = txtUsername.Text
            Pwd = txtPassword.Text
            'Dim frm2 As New Form2
            Const LOGON32_LOGON_INTERACTIVE As Long = 2
            Const LOGON32_LOGON_NETWORK As Long = 3
            Const LOGON32_PROVIDER_DEFAULT As Long = 0
            Const LOGON32_PROVIDER_WINNT50 As Long = 3
            Const LOGON32_PROVIDER_WINNT40 As Long = 2
            Const LOGON32_PROVIDER_WINNT35 As Long = 1


            tokenHandle = IntPtr.Zero
            'Call the LogonUser function to obtain a handle to an access token.
            Dim returnValue As Boolean = LogonUser(UserName, MachineName, Pwd, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, tokenHandle)

            If returnValue = False Then
                'This function returns the error code that the last unmanaged function returned.
                Dim ret As Integer = Marshal.GetLastWin32Error()
                Dim errmsg As String = GetErrorMessage(ret)
                lblErrorMessage.Text = errmsg.Replace(":", ": ")
            Else
                'Create the WindowsIdentity object for the Windows user account that is
                'represented by the tokenHandle token.
                Dim newId As New WindowsIdentity(tokenHandle)
                Dim userperm As New WindowsPrincipal(newId)
                'Verify whether the Windows user has administrative credentials.

                lblErrorMessage.Text = "Login success: " & userperm.Identity.Name() & " has been successfully authenticated."
                'If userperm.IsInRole(WindowsBuiltInRole.Administrator) Then
                '    MsgBox(userperm.Identity.Name())
                '    MsgBox(userperm.Identity.AuthenticationType)
                '    MsgBox(userperm.Identity.IsAuthenticated)

                'Else

                'End If
            End If

            'Free the access token.
            If Not System.IntPtr.op_Equality(tokenHandle, IntPtr.Zero) Then
                CloseHandle(tokenHandle)
            End If
        Catch ex As Exception
            Error_Handler(ex)

        End Try
        Try

            'Free the access token.
            If Not System.IntPtr.op_Equality(tokenHandle, IntPtr.Zero) Then
                CloseHandle(tokenHandle)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                If identifier_msg = "" Then
                    Dim Display_Message1 As New Display_Message("The following application error has occurred: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString() & vbCrLf & "The application will attempt to recover from this shortly.")
                    Display_Message1.ShowDialog()
                Else
                    Dim Display_Message1 As New Display_Message("The following application error has occurred: " & vbCrLf & ex.Message.ToString() & vbCrLf & "The application will attempt to recover from this shortly.")
                    Display_Message1.ShowDialog()
                End If
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                If identifier_msg = "" Then
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & ex.ToString)
                Else
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                End If

                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Windows User Authentication's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub
End Class
