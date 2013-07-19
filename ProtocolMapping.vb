' Name:     Protocol Mapping
' Author:   Donald G Plugge
' Date:     12/8/2010
' Purpose:  Class to mapping the protocol headers with parameter names

Public Class ProtocolMapping

    Private mProtocolHeader As ArrayList
    Private mParNames As ArrayList
    Private mHeaderMap As Hashtable
    Private mParMap As Hashtable

    ' dgp rev 12/8/2010 
    Public Sub New()

        mProtocolHeader = New ArrayList
        mParNames = New ArrayList
        mHeaderMap = New Hashtable
        mParMap = New Hashtable

    End Sub

    Public Function ReMap(ByVal col As Integer, ByVal par As Integer) As Boolean

        ReMap = True

        If col < 0 Or col >= mProtocolHeader.Count Then Return False
        If par < 0 Or par >= mParNames.Count Then Return False

        If CrossIndex.Count = 0 Then
            mCrossIndex.Add(col, par)
            Return True
        End If

        Dim item As DictionaryEntry

        Dim tmp As Hashtable = New Hashtable
        tmp.Add(col, par)

        For Each item In CrossIndex
            If Not (item.Key = col Or item.Value = par) Then
                tmp.Add(item.Key, item.Value)
            End If
        Next

        mCrossIndex = tmp
        Return True

    End Function



    Public Function ReMapx(ByVal col As Integer, ByVal par As Integer) As Boolean


        If col < 0 Or col >= mProtocolHeader.Count Then Return False
        If par < 0 Or par >= mParNames.Count Then Return False

        If CrossIndex.Count = 0 Then
            mCrossIndex.Add(col, par)
            Return True
        End If

        Dim item As DictionaryEntry
        Dim ReverseIndex As New Hashtable
        For Each item In mCrossIndex
            ReverseIndex.Add(item.Value, item.Key)
        Next

        If ReverseIndex.ContainsKey(par) Then
            ' par match
            If CrossIndex.ContainsKey(col) Then
                ' par and col match
                ' null action, already mappped
                If CrossIndex(col) = par Then Exit Function
                ' par match only
            End If
            CrossIndex.Remove(ReverseIndex(par))
        Else
            If CrossIndex.ContainsKey(col) Then
                CrossIndex.Remove(col)
            End If
        End If
        CrossIndex.Add(col, par)

    End Function

    ' dgp rev 12/8/2010 
    Public Sub Header(ByVal HeadItems As ArrayList)

        If HeadItems Is Nothing Then Return
        If HeadItems.Count = 0 Then Return
        mProtocolHeader = HeadItems
        mHeaderMap = New Hashtable

        Dim item
        Dim idx As Integer = 0
        Dim match As String
        For Each item In mProtocolHeader
            If FCSAntibodies.WillMatch(item) Then
                match = FCSAntibodies.FindMatch(item)
                ' dgp rev 1/12/2011 problem is two matches, so run a check
                If Not mHeaderMap.ContainsKey(match) Then mHeaderMap.Add(FCSAntibodies.FindMatch(item), idx)
            End If
            idx += 1
        Next

    End Sub

    ' dgp rev 12/8/2010 
    Private mValid As Boolean = False
    Public ReadOnly Property Valid As Boolean
        Get
            Return (mParNames IsNot Nothing And mProtocolHeader IsNot Nothing)
        End Get
    End Property


    Private mCrossIndex As Hashtable = Nothing
    Public ReadOnly Property CrossIndex As Hashtable
        Get
            If mCrossIndex Is Nothing Then
                mCrossIndex = New Hashtable
                If mParMap Is Nothing Then Return mCrossIndex
                If mHeaderMap Is Nothing Then Return mCrossIndex
                Dim item As DictionaryEntry
                For Each item In mParMap
                    If mHeaderMap.ContainsKey(item.Key) Then
                        mCrossIndex.Add(mHeaderMap(item.Key), item.Value)
                    End If
                Next
            End If
            Return mCrossIndex
        End Get
    End Property

    ' dgp rev 12/8/2010 
    Public ReadOnly Property ParMap As Hashtable
        Get
            Return mParMap
        End Get
    End Property

    ' dgp rev 12/8/2010 
    Public ReadOnly Property HeaderMap As Hashtable
        Get
            Return mHeaderMap
        End Get
    End Property

    ' dgp rev 12/8/2010 
    Public Sub ParNames(ByVal ParItems As ArrayList)

        If ParItems Is Nothing Then Return
        If ParItems.Count = 0 Then Return
        mParNames = ParItems
        mParMap = New Hashtable

        Dim item
        Dim idx As Integer = 0
        For Each item In mParNames
            Dim key = FCSAntibodies.FindMatch(item)
            If FCSAntibodies.WillMatch(item) Then
                If Not mParMap.ContainsKey(key) Then mParMap.Add(key, idx)
            End If
            idx += 1
        Next

    End Sub


End Class
