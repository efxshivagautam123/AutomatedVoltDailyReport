Module modSharedItems
    'Test
    'Public strSQLPassword As String = "Edgew!sE"
    'Public strSQLServer As String = "eedc02"
    'Public strSQLUser As String = "edgewise"
    'Public strSQLDatabase As String = "Edgewise"

    ''Verticomm SQL1 - Production Database
    'Public strSQLServer As String = "EE-SQL01"
    'Public strSQLUser As String = "edgewise3"
    'Public strConnectionString As String = _
    '"server = EE-SQL01;User ID=edgewise3;Password=Edgew!sE;" & _
    '"database=edgewise"
    'Public objConnection As New SqlConnection(strConnectionString)

    'Verticomm SQL2 - Reports/Wages
    Public strSQLServer As String = "EE-SQL02"
    Public strSQLUser As String = "edgewise3"
    Public strConnectionString As String = _
    "server = EE-SQL02;User ID=edgewise3;Password=Edgew!sE;" & _
    "database=Edgewise"
    Public objConnection As New SqlConnection(strConnectionString)

    ''SIDES Test Database!
    'Public strSQLServer As String = "EE-SQL01"
    'Public strSQLUser As String = "edgewise3"
    'Public strConnectionString As String = _
    '"server = EE-SQL01;User ID=edgewise3;Password=Edgew!sE;" & _
    '"database=edgewisesidestest"
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
