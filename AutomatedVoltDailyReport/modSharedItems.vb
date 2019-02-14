Module modSharedItems

    'Verticomm SQL2 - Reports/Wages
    Public strSQLServer As String = "EE-SQL02"
    Public strSQLUser As String = "edgewise3"
    Public strConnectionString As String =
    "server = EE-SQL02;User ID=edgewise3;Password=Edgew!sE;" &
    "database=Edgewise"
    Public objConnection As New SqlConnection(strConnectionString)



    ''Verticomm SQL3 - Future Production Reports/Wages Database
    'Public strSQLServerRpt As String = "EE-SQL03"
    'Public strSQLUserRpt As String = "edgewise3"
    'Public strConnectionString As String =
    '"server = EE-SQL03;User ID=edgewise3;Password=Edgew!sE;" &
    '"database=edgewiserpt"
    'Public objConnection As New SqlConnection(strConnectionString)


    Public bolAttribAdministrator As Boolean
    Public bolAttribManager As Boolean
    Public strUsername As String
    Public intPublicClientID As Integer

    Public Function IsUseableDate(ByVal strDateIn As String) As Boolean
        Dim bolIsDate As Boolean = False

        If IsDate(strDateIn) Then
            If CDate(strDateIn) < #1/1/1900# Or CDate(strDateIn) > #1/1/2100# Then
                bolIsDate = False
            Else
                bolIsDate = True
            End If
        End If

        Return bolIsDate
    End Function
End Module
