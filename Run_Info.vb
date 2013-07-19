' author:  Donald G Plugge
' date:    6/5/08
' purpose: Tracking run information

Public Class Run_Info

    'Declare data members
    Private m_Run_Name As String
    Public Property Run_Name() As String
        Get
            Return m_Run_Name
        End Get
        Set(ByVal value As String)
            m_Run_Name = value
        End Set
    End Property

    Private m_Valid As Boolean
    Public Property valid() As Boolean
        Get
            Return m_Valid
        End Get
        Set(ByVal value As Boolean)
            m_Valid = value
        End Set
    End Property
    ' dgp rev 5/31/07 Files in Run
    Private m_Files As Collection
    Public Property Files() As Collection
        Get
            Return m_Files
        End Get
        Set(ByVal value As Collection)
            m_Files = value
        End Set
    End Property

    Private m_Missing As Collection
    Public Property Missing() As Collection
        Get
            Return m_Missing
        End Get
        Set(ByVal value As Collection)
            m_Missing = value
        End Set
    End Property

End Class
