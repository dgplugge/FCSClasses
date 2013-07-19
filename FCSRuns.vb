Imports FCS_Classes

Public Class FCSRuns


    Public Shared Function MachineDateMatch(ByVal run) As ArrayList

        Dim source_run As Run_Name = New Run_Name(run)
        MachineDateMatch = New ArrayList

        Dim testrun As Run_Name
        Dim eachrun
        For Each eachrun In System.IO.Directory.GetDirectories(FlowStructure.Data_Root)

            testrun = New Run_Name(System.IO.Path.GetFileNameWithoutExtension(eachrun))
            If testrun.machine = source_run.Machine Then
                If testrun.NormDate = source_run.NormDate Then
                    MachineDateMatch.Add(System.IO.Path.GetFileNameWithoutExtension(eachrun))
                End If
            End If

        Next

    End Function


End Class
