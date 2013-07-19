' Name: FCSConvert Class
' Author: Donald G Plugge
' Date: 12/11/06 
' Purpose: Class used to facilitate the conversion of FCS data

Public Class FCSConvert
    ' dgp rev 12/1806 enumerate the negative handling options.
    Public Enum NegEnum
        allow = 1
        clip = 2
        zero = 3
    End Enum

    Private m_flg_Run As Boolean = False
    Private m_flg_Table As Boolean = False

    ' dgp rev 12/18/06 Negative number handling
    Private m_neg_mode As NegEnum = NegEnum.clip
    Public Property Neg_Mode() As NegEnum
        Get
            Return m_neg_mode
        End Get
        Set(ByVal value As NegEnum)
            m_neg_mode = value
        End Set
    End Property
    ' dgp rev 12/18/06 structure describing individual parameters
    Private Structure m_param
        Public name As String
        Public index As Integer
        Public log As Boolean
        Public reso As Integer
    End Structure

    ' dgp rev 12/15/06 a collection of information relative to individual 
    ' parameters log/lin, resolution, name
    Private m_params_info As Collection
    Public Property Params_Info() As Collection
        Get
            Return m_params_info
        End Get
        Set(ByVal value As Collection)
            m_params_info = value
        End Set
    End Property
    ' list of parameters to convert
    Private m_param_idx As ArrayList
    Public Property Param_Idx() As ArrayList
        Get
            Return m_param_idx
        End Get
        Set(ByVal value As ArrayList)
            m_param_idx = value
        End Set
    End Property
    ' dgp rev 12/18/06 FCS run to convert
    Private m_Run As FCSRun
    Public Property Convert_Run() As FCSRun
        Get
            Return m_Run
        End Get
        Set(ByVal value As FCSRun)
            m_Run = value
        End Set
    End Property
    ' dgp rev 12/18/06 must have an FCS file to convert
    Private m_FCS_File As FCS_File
    Public Property FCS_Source() As FCS_File
        Get
            Return m_FCS_File
        End Get
        Set(ByVal value As FCS_File)
            m_FCS_File = value
        End Set
    End Property
    ' dgp rev 12/18/06 an XML table may be merge into the dataset
    Private m_table As FCSTable
    Public Property Table() As FCSTable
        Get
            Return m_table
        End Get
        Set(ByVal value As FCSTable)
            m_table = value
        End Set
    End Property

End Class
