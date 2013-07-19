' Author: Donald G Plugge
' Date:   5/22/08
' Name:   Run Compare
Public Class RunCompare

    Implements IComparer

    Public Shared Idx As Int16 = 3
    Public Shared Reps As Integer

    Dim xrun As Run_Name
    Dim yrun As Run_Name
    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function UserComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        ' dgp rev 3/28/09 Watch for blank dates - xyzzy
        If (xrun.User = "" Or yrun.User = "") Then
            If (xrun.User = "") Then Return 0
            Return 1
        Else
            Try
                If (xrun.User.ToLower = yrun.User.ToLower) Then
                    Return Date.Compare(xrun.NormDate, yrun.NormDate)
                Else
                    Return String.Compare(xrun.User, yrun.User)
                End If
            Catch ex As Exception
            End Try
        End If

    End Function

    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function RunComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        ' dgp rev 3/28/09 Watch for blank dates - xyzzy
        If (xrun.RunName = "" Or yrun.RunName = "") Then
            If (xrun.RunName = "") Then Return 0
            Return 1
        Else
            Try
                If (xrun.RunNum = yrun.RunNum) Then
                    Return Date.Compare(xrun.NormDate, yrun.NormDate)
                Else
                    Return (xrun.RunNum > yrun.RunNum)
                End If
            Catch ex As Exception
            End Try
        End If

    End Function

    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function MachineComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        ' dgp rev 3/28/09 Watch for blank dates - xyzzy
        If (xrun.Machine = "" Or yrun.Machine = "") Then
            If (xrun.Machine = "") Then Return 0
            Return 1
        Else
            Try
                If (xrun.Machine.ToLower = yrun.Machine.ToLower) Then
                    Return Date.Compare(xrun.NormDate, yrun.NormDate)
                Else
                    Return String.Compare(xrun.Machine, yrun.Machine)
                End If
            Catch ex As Exception
            End Try
        End If

    End Function

    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function RecentComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        Dim xdate As DateTime
        Dim ydate As DateTime

        xdate = System.IO.File.GetLastWriteTime((System.IO.Path.Combine(FlowStructure.Data_Root, x)))
        ydate = System.IO.File.GetLastWriteTime((System.IO.Path.Combine(FlowStructure.Data_Root, y)))

        Try
            If (xdate = ydate) Then
                Return String.Compare(xrun.User, yrun.User)
            Else
                Return Date.Compare(xdate, ydate)
            End If
        Catch ex As Exception
        End Try

    End Function


    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function DateComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        ' dgp rev 3/28/09 Watch for blank dates - xyzzy
        If (xrun.Dat = "" Or yrun.Dat = "") Then
            If (xrun.Dat Is Nothing) Then Return 1
            If (yrun.Dat Is Nothing) Then Return -1
        Else
            Try
                If (xrun.Dat.ToLower = yrun.Dat.ToLower) Then
                    Return String.Compare(xrun.User, yrun.User)
                Else
                    Return Date.Compare(xrun.NormDate, yrun.NormDate)
                End If
            Catch ex As Exception
            End Try
        End If

    End Function

    ' dgp rev 5/22/08 Compare based upon User Name
    Private Function DateOnlyComp(ByVal x As Object, ByVal y As Object) As Integer

        xrun = New Run_Name(x)
        xrun.ParseRun()
        yrun = New Run_Name(y)
        yrun.ParseRun()

        ' dgp rev 3/28/09 Watch for blank dates - xyzzy
        If (xrun.Dat = "" Or yrun.Dat = "") Then
            If (xrun.Dat Is Nothing) Then Return 1
            If (yrun.Dat Is Nothing) Then Return -1
        Else
            Try
                Return Date.Compare(xrun.NormDate, yrun.NormDate)
            Catch ex As Exception
            End Try
        End If

    End Function

    ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
       Implements IComparer.Compare

        Reps = Reps + 1

        Select Case Idx
            Case 1
                Return RunComp(x, y)
            Case 2
                Return UserComp(x, y)
            Case 3
                Return DateComp(x, y)
            Case 4
                Return MachineComp(x, y)
            Case 5
                Return RecentComp(x, y)
            Case 6
                Return DateOnlyComp(x, y)
        End Select

    End Function 'IComparer.Compare


End Class
