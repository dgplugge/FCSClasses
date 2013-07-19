' dgp rev 6/6/07 FCS run name
Public Class Run_Name

    Private m_Full_Name As String

    Private m_VMS_Flag As Boolean = False
    Private m_Run_Flag As Boolean = False
    Private m_Pos As Int16

    Private Function IsMachine(ByVal item As String) As Boolean

        Return (Char.IsLetter(item.Substring(0, 1)) And Char.IsDigit(item.Substring(item.Length - 1, 1)))

    End Function

    Private Function IsTime(ByVal item As String) As Boolean

        Return (Char.IsDigit(item.Substring(0, 1)) And item.Contains("!"))

    End Function

    Private Function IsRun(ByVal item As String) As Boolean

        If (item.Length < 6) Then Return False ' not long enough

        Return (item.Substring(0, 1).ToLower = "r" And Char.IsDigit(item.Substring(1, 5)))

    End Function

    Private Function IsUser(ByVal item As String) As Boolean

        Return (Char.IsLetter(item))

    End Function



    Private Function IsDate(ByVal item As String) As Boolean

        Return (item.Contains("-"))

    End Function

    Private Sub Catagorize(ByVal Item As String)

        While (m_Pos < 5)
            Select Case m_Pos
                Case 0
                    m_Pos = m_Pos + 1
                    m_Machine = Item
                    Exit Sub
                Case 1
                    m_Pos = m_Pos + 1
                    If (IsDate(Item)) Then
                        m_Date = Item
                        Exit Sub
                    End If
                Case 2
                    m_Pos = m_Pos + 1
                    If (IsTime(Item)) Then
                        m_Time = Item
                        Exit Sub
                    End If
                Case 3
                    m_Pos = m_Pos + 1
                    If (IsUser(Item)) Then
                        m_User = Item
                        Exit Sub
                    End If
                Case 4
                    m_Pos = m_Pos + 1
                    If (IsRun(Item)) Then
                        m_RunName = Item
                        Exit Sub
                    End If
            End Select
        End While

    End Sub

    ' dgp rev 5/19/09 catagorize each item
    Public Sub ParseRun()

        Dim item
        m_Pos = 0

        For Each item In m_arr
            Catagorize(item)
        Next

    End Sub

    Private m_Machine As String

    Public Enum CompStat
        [Match] = 1
        Nomatch = -1
        Mismatch = -2
    End Enum

    Public ReadOnly Property Machine() As String
        Get
            Return m_Machine
        End Get
    End Property

    Private m_NormDate As DateTime
    Public ReadOnly Property NormDate() As DateTime
        Get
            If (m_Date = "") Then Return ""
            If (m_Date.Length > 9) Then
                m_NormDate = DateTime.ParseExact(m_Date, "dd-MMM-yyyy", Nothing)
            Else
                m_NormDate = DateTime.ParseExact(m_Date, "dd-MMM-yy", Nothing)
            End If
            Return m_NormDate
        End Get
    End Property

    Private m_Date As String
    Public ReadOnly Property Dat() As String
        Get
            Return m_Date
        End Get
    End Property
    Private m_Time As String
    Public ReadOnly Property Time() As String
        Get
            Return m_Time
        End Get
    End Property
    Private m_User As String
    Public ReadOnly Property User() As String
        Get
            Return m_User
        End Get
    End Property
    Private m_RunName As String
    Public ReadOnly Property RunName() As String
        Get
            Return m_RunName
        End Get
    End Property
    Public ReadOnly Property RunNum() As Integer
        Get
            Return CInt(m_RunName.ToLower.Replace("r", ""))
        End Get
    End Property

    Private m_arr() As String

    Public ReadOnly Property MDTUR() As String
        Get
            If Not mMDT_Flag Then Return ""
            If mAssigned Then
                Return m_Machine + "_" + m_Date + "_" + m_Time + "_" + m_User + "_" + m_RunName
            Else
                Return m_Machine + "_" + m_Date + "_" + m_Time
            End If
        End Get
    End Property

    Private mAssigned As Boolean = False
    Private mMDT_Flag As Boolean = False
    Public ReadOnly Property MDT_Flag() As Boolean
        Get
            Return mMDT_Flag
        End Get
    End Property

    Private mFormatted As Boolean = False
    ' dgp rev 5/19/09 
    Public Sub New(ByVal Run As String)

        m_arr = Run.ToLower.Split("_")

        ' dgp rev 5/20/09 check the two valid formats
        Select Case m_arr.Length
            Case 3
                If (Me.IsDate(m_arr(1))) Then
                    m_Machine = m_arr(0)
                    m_Date = m_arr(1)
                    m_Time = m_arr(2)
                    mFormatted = True
                End If
            Case 5
                If (Me.IsDate(m_arr(1)) And Me.IsRun(m_arr(4))) Then
                    m_Machine = m_arr(0)
                    m_Date = m_arr(1)
                    m_Time = m_arr(2)
                    m_User = m_arr(3)
                    m_RunName = m_arr(4)
                    mFormatted = True
                End If
        End Select

        mAssigned = (Not (m_User Is Nothing Or m_RunName Is Nothing))
        mMDT_Flag = (Not (m_Machine Is Nothing Or m_Date Is Nothing Or m_Time Is Nothing))

    End Sub

    Public Function Compare(ByVal Test As String) As CompStat

        Dim t_arr() As String = Test.ToLower.Split("_")
        Dim t_run_flag As Boolean
        Dim t_match As Integer
        Dim t_run_match As Boolean = False

        Compare = CompStat.Nomatch

        ' minimal requirements
        If (t_arr.Length > 2) Then
            t_run_flag = True
            ' non User/Run format
            If (t_arr.Length = 3) Then
                t_match = (t_arr(0) = m_Machine)
                t_match = t_match + (t_arr(1) = m_Date)
                t_match = t_match + (t_arr(2) = m_Time)
                If (t_match = -3) Then Compare = CompStat.Match
                ' User/Run format
            ElseIf (m_arr.Length = 5) Then
                m_VMS_Flag = True
                t_match = 0
                t_match = t_match + (t_arr(3) = m_User)
                t_match = t_match + (t_arr(4) = m_RunName)
                ' User/Run match
                If (t_match = -2) Then t_run_match = True
                t_match = t_match + (t_arr(0) = m_Machine)
                t_match = t_match + (t_arr(1) = m_Date)
                t_match = t_match + (t_arr(2) = m_Time)
                If (t_match = -5) Then
                    ' Full match
                    Compare = CompStat.Match
                Else
                    ' No match, but perhaps partial match on User/Run
                    If (t_run_match) Then Compare = CompStat.Mismatch
                End If
            End If
        End If

    End Function

End Class
