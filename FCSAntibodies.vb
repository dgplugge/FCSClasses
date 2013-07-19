Imports System.Text.RegularExpressions

Public Class FCSAntibodies

    'dgp rev 8/5/2010 
    Public Shared Function WillMatch(ByVal text As String) As Boolean

        text = text.ToLower

        If Regex.IsMatch(text, "pe") Then Return True
        If Regex.IsMatch(text, "apc") Then Return True
        If Regex.IsMatch(text, "fitc") Then Return True
        If Regex.IsMatch(text, "(488|647|594|670)") Then Return True
        If Regex.IsMatch(text, ".*cy.*[0-9]") Then Return True
        If Regex.IsMatch(text, "(p.*b.*)") Then Return True

        Return False

    End Function



    'dgp rev 8/5/2010 
    Public Shared Function FindMatch(ByVal orig As String) As String

        Dim text = orig.ToLower.Replace("#", "")
        Dim results As MatchCollection

        If Regex.IsMatch(text, "^pe") Then
            If Regex.IsMatch(text, ".*cy.*[0-9]") Then
                results = Regex.Matches(text, ".*cy.*([0-9])")
                Return String.Format("PE-Cy{0}", results.Item(0).Groups(results.Item(0).Groups.Count - 1))
            Else
                Return "PE"
            End If
        ElseIf Regex.IsMatch(text, "apc") Then
            If Regex.IsMatch(text, ".*cy.*[0-9]") Then
                results = Regex.Matches(text, ".*cy.*([0-9])")
                Return String.Format("APC-Cy{0}", results.Item(0).Groups(results.Item(0).Groups.Count - 1))
            Else
                Return "APC"
            End If
        ElseIf Regex.IsMatch(text, "fitc") Then
            Return "FITC"
        ElseIf Regex.IsMatch(text, "(488|647|594|670)") Then
            results = Regex.Matches(text, "(488|647|594|670)")
            Return String.Format("Alexa {0}", results.Item(0))
        ElseIf Regex.IsMatch(text, ".*cy.*[0-9]") Then
            results = Regex.Matches(text, ".*cy.*([0-9])")
            Return String.Format("Cy{0}", results.Item(results.Count - 1))
        ElseIf Regex.IsMatch(text, "(p.*b.*)") Then
            results = Regex.Matches(text, "(p.*b.*)")
            Return String.Format("Pacific Blue")
        End If

        Return orig.Replace("#", "")

    End Function



End Class
