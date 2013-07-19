Imports System
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Security.Permissions
Imports Microsoft.VisualBasic
<Assembly: SecurityPermissionAttribute(SecurityAction.RequestMinimum, UnmanagedCode:=True), _
 Assembly: PermissionSetAttribute(SecurityAction.RequestMinimum, Name:="FullTrust")> 

Public Class RunAs_Impersonator
#Region "Private Variables and Enum Constants"
    Private tokenHandle As New IntPtr(0)
    Private dupeTokenHandle As New IntPtr(0)
    Private impersonatedUser As WindowsImpersonationContext

    Private m_username As String
    Public ReadOnly Property Valid_User() As String
        Get
            Return m_username
        End Get
    End Property
    Private m_domain As String
    Public ReadOnly Property Valid_Domain() As String
        Get
            Return m_domain
        End Get
    End Property
    Private m_password As String
    Private m_valid As Boolean = False
    Public ReadOnly Property Valid() As Boolean
        Get
            Return m_valid
        End Get
    End Property

#End Region
#Region "Properties"

#End Region
#Region "Public Methods"
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean

    Public Declare Auto Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As IntPtr, _
      ByVal SECURITY_IMPERSONATION_LEVEL As Integer, _
      ByRef DuplicateTokenHandle As IntPtr) As Boolean

    ' Test harness.
    ' If you incorporate this code into a DLL, be sure to demand FullTrust.
    <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
    Public Sub ImpersonateStart(ByVal Domain As String, ByVal userName As String, ByVal Password As String)
        Try
            tokenHandle = IntPtr.Zero
            ' Call LogonUser to obtain a handle to an access token.
            Dim returnValue As Boolean = LogonUser(userName, Domain, Password, 2, 0, tokenHandle)

            'check if logon successful
            If returnValue = False Then
                Dim ret As Integer = Marshal.GetLastWin32Error()
                Console.WriteLine("LogonUser failed with error code : {0}", ret)
                Throw New System.ComponentModel.Win32Exception(ret)
                Exit Sub
            End If

            'Logon succeeded

            ' Use the token handle returned by LogonUser.
            Dim newId As New WindowsIdentity(tokenHandle)
            impersonatedUser = newId.Impersonate()
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try
        m_username = userName
        m_domain = Domain
        m_password = Password
        m_valid = True
        'MsgBox("running as " & impersonatedUser.ToString & " -- " & WindowsIdentity.GetCurrent.Name)
    End Sub
    <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
    Public Sub ImpersonateStop()
        ' Stop impersonating the user.
        If (Not impersonatedUser Is Nothing) Then impersonatedUser.Undo()

        ' Free the tokens.
        If Not System.IntPtr.op_Equality(tokenHandle, IntPtr.Zero) Then
            CloseHandle(tokenHandle)
        End If
        'MsgBox("running as " & Environment.UserName)
    End Sub
#End Region
#Region "Private Methods"
    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
     ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
     ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
     ByRef phToken As IntPtr) As Boolean

    <DllImport("kernel32.dll")> _
    Public Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, _
     ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], _
     ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function
#End Region
End Class
