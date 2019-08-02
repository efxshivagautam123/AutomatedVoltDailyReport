Module modSharedItems
    Public bolV4Security As Boolean = True
    Public strSQLDatabase As String = ""
    Public strDevSQLDatabase As String = ""
    Public strSQLServer As String = ""
    Public strSQLUser As String = ""
    Public strSQLPassword As String = ""
    Public strSQLServerRpt As String = ""
    Public strSQLUserRpt As String = ""


    'SQL03 Production ****************************************************************************

    Public DatabaseEnvironment As String = "Production"

    'Verticomm SQL3 - Production Claims Database 
    Public strConnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("ProductionConnectionString").ConnectionString
    Public objConnection As New SqlConnection(strConnectionString)


    'Verticomm SQL3 - Reports/Wage Database 
    Public strConnectionStringRpt As String = System.Configuration.ConfigurationManager.ConnectionStrings("ReportConnectionString").ConnectionString
    Public objConnectionRpt As New SqlConnection(strConnectionStringRpt)


    'Verticomm SQL3 - Dev Database (used in wage file testing within production)
    Public strConnectionStringDev As String = System.Configuration.ConfigurationManager.ConnectionStrings("DevConnectionString").ConnectionString
    Public objConnectionDev As New SqlConnection(strConnectionStringRpt)

    '***************************************************************************************




    ''SQL03 Dev ********************************************************************

    'Public DatabaseEnvironment As String = "Dev"

    ''Verticomm SQL3 - Dev Database 
    'Public strConnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("DevConnectionString").ConnectionString
    'Public objConnection As New SqlConnection(strConnectionString)


    ''Verticomm SQL3 - Dev Database 
    'Public strConnectionStringRpt As String = System.Configuration.ConfigurationManager.ConnectionStrings("DevConnectionString").ConnectionString
    'Public objConnectionRpt As New SqlConnection(strConnectionStringRpt)

    ''Verticomm SQL3 - Dev Database (used In wage file testing within production)
    'Public strConnectionStringDev As String = System.Configuration.ConfigurationManager.ConnectionStrings("DevConnectionString").ConnectionString
    'Public objConnectionDev As New SqlConnection(strConnectionStringRpt)

    ''***************************************************************************************


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
