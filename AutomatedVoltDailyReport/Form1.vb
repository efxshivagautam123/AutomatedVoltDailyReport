Imports System
Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Security.Principal
Imports WinSCP
Public Class Form1
    Dim strRecord As String
    Dim strFileName As String
    Dim strShortFileName As String
    Dim FileNum As Integer = 0
    Dim strDateToday As String = Format(DateTime.Today, "yyyyMMdd")
    Dim strDateSixMonthsPast As String = Format(DateAdd(DateInterval.Day, 1, DateAdd(DateInterval.Month, -6, DateTime.Today)), "yyyMMdd")
    Dim x As Integer
    Dim strTemp As String
    Dim strSSNIn As String
    Dim strPreviousSSN As String = "x"
    Dim cmdSQL As New SqlCommand
    Dim cmdSQL2 As New SqlCommand
    Dim objConnection2 As New SqlConnection(strConnectionString)
    Dim intHldClientID As Integer = 254  'Volt 115500
    Dim strHldFEIN As String
    Dim bolInHouse As Boolean


    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim ident As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim user As New WindowsPrincipal(ident)
        Dim strUsername As String = user.Identity.Name
        'Dim strShortUsername As String = Mid(strUsername, 9, Len(strUsername) - 8)
        Dim strShortUsername As String = Mid(strUsername, 4, Len(strUsername) - 3)     'EE domain - Verticomm 2018
        Dim x As Integer = -1

        'Check to see if administrator
        cmdSQL.CommandText = "select count(*) from UserAttributes att " & _
            "inner join Users us on us.intUserID = att.intUserID " & _
            "inner join SecurityAttributes sec on sec.intAttributeID = att.intAttributeID " & _
            "where sec.strAttribute = 'Administrator' and us.strShortUserName = '" & strShortUsername & "'"
        cmdSQL.Connection = objConnection
        objConnection.Open()
        x = cmdSQL.ExecuteScalar()
        objConnection.Close()


        If x < 1 Then
            MessageBox.Show("You are not authorized for this screen!")
            End
        Else
            'Administrator - Proceed
        End If

        Dim path = CreateDailyReport()
        UploadReport(path)
        Archive(path)
        File.Delete(path)

        Me.Close()
    End Sub

    Private Function CreateDailyReport() As String
        Dim strBegWageLoc As String
        Dim strFullWageLoc As String
        Dim strMidWageLoc As String
        Dim strEndWageLoc As String
        Dim strNewDeptID As String

        'Find Appropriate Filename
        If Dir("C:\Volt\Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & ".txt") = "" Then
            strFileName = "C:\Volt\Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & ".txt"
            strShortFileName = "Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & ".txt"
        Else
            Do
                FileNum = FileNum + 1
                If Dir("C:\Volt\Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & "_" & CStr(FileNum) & ".txt") = "" Then
                    strFileName = "C:\Volt\Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & "_" & CStr(FileNum) & ".txt"
                    strShortFileName = "Voltrpt_mthly_" & strDateSixMonthsPast & "-" & strDateToday & "_" & CStr(FileNum) & ".txt"
                    Exit Do
                End If
            Loop
        End If

        Dim DailyFile As New System.IO.StreamWriter(strFileName)

        cmdSQL.CommandType = CommandType.StoredProcedure
        cmdSQL.CommandText = "spVoltDailyReport_v2"
        cmdSQL.CommandTimeout = 10800
        cmdSQL.Connection = objConnection
        objConnection.Open()

        Dim DailyReader As SqlDataReader = cmdSQL.ExecuteReader()

        While DailyReader.Read()
            bolInHouse = False
            strRecord = DailyReader(0)
            strSSNIn = Mid(strRecord, 2, 9)

            strFullWageLoc = Mid(strRecord, 27, 9)
            strBegWageLoc = Mid(strRecord, 27, 3)
            strMidWageLoc = Mid(strRecord, 30, 3)
            strEndWageLoc = Mid(strRecord, 33, 3)

            'Check for 118 or 218 RDC Dept ID
            Select Case strBegWageLoc
                Case "116", "118", "216", "218"
                    Select Case strMidWageLoc
                        Case "002"
                            strNewDeptID = "102001" & strEndWageLoc
                        Case "003"
                            strNewDeptID = "102004" & strEndWageLoc
                        Case "004"
                            strNewDeptID = "102015" & strEndWageLoc
                        Case "005"
                            strNewDeptID = "102016" & strEndWageLoc
                        Case "006"
                            strNewDeptID = "102017" & strEndWageLoc
                        Case "007"
                            strNewDeptID = "102018" & strEndWageLoc
                        Case "008"
                            strNewDeptID = "102019" & strEndWageLoc
                        Case "009"
                            strNewDeptID = "102033" & strEndWageLoc
                        Case "010"
                            strNewDeptID = "102035" & strEndWageLoc
                        Case "011"
                            strNewDeptID = "102036" & strEndWageLoc
                        Case "012"
                            strNewDeptID = "102045" & strEndWageLoc
                        Case "013"
                            strNewDeptID = "102054" & strEndWageLoc
                        Case "014"
                            strNewDeptID = "102057" & strEndWageLoc
                        Case "015"
                            strNewDeptID = "102059" & strEndWageLoc
                        Case "016"
                            strNewDeptID = "102061" & strEndWageLoc
                        Case "017"
                            strNewDeptID = "102066" & strEndWageLoc
                        Case "018"
                            strNewDeptID = "102068" & strEndWageLoc
                        Case "019"
                            strNewDeptID = "102076" & strEndWageLoc
                        Case "020"
                            strNewDeptID = "102086" & strEndWageLoc
                        Case "021"
                            strNewDeptID = "102084" & strEndWageLoc
                        Case "022"
                            strNewDeptID = "102098" & strEndWageLoc
                        Case "025"
                            strNewDeptID = "102180" & strEndWageLoc
                        Case "026"
                            strNewDeptID = "191027" & strEndWageLoc
                        Case "027"
                            strNewDeptID = "102191" & strEndWageLoc
                        Case "028"
                            strNewDeptID = "102603" & strEndWageLoc
                        Case "029"
                            strNewDeptID = "102417" & strEndWageLoc
                        Case "030"
                            strNewDeptID = "102424" & strEndWageLoc
                        Case "031"
                            strNewDeptID = "105085" & strEndWageLoc
                        Case "032"
                            strNewDeptID = "102023" & strEndWageLoc
                        Case "033"
                            strNewDeptID = "102447" & strEndWageLoc
                        Case "034"
                            strNewDeptID = "102448" & strEndWageLoc
                        Case "035"
                            strNewDeptID = "102449" & strEndWageLoc
                        Case "036"
                            strNewDeptID = "102456" & strEndWageLoc
                        Case "037"
                            strNewDeptID = "102458" & strEndWageLoc
                        Case "038"
                            strNewDeptID = "116001" & strEndWageLoc
                        Case "039"
                            strNewDeptID = "104004" & strEndWageLoc
                        Case "040"
                            strNewDeptID = "104030" & strEndWageLoc
                        Case "041"
                            strNewDeptID = "104033" & strEndWageLoc
                        Case "042"
                            strNewDeptID = "104034" & strEndWageLoc
                        Case "043"
                            strNewDeptID = "104035" & strEndWageLoc
                        Case "044"
                            strNewDeptID = "104039" & strEndWageLoc
                        Case "045"
                            strNewDeptID = "104041" & strEndWageLoc
                        Case "046"
                            strNewDeptID = "104042" & strEndWageLoc
                        Case "047"
                            strNewDeptID = "104060" & strEndWageLoc
                        Case "048"
                            strNewDeptID = "104080" & strEndWageLoc
                        Case "049"
                            strNewDeptID = "104083" & strEndWageLoc
                        Case "050"
                            strNewDeptID = "104086" & strEndWageLoc
                        Case "051"
                            strNewDeptID = "104092" & strEndWageLoc
                        Case "052"
                            strNewDeptID = "104093" & strEndWageLoc
                        Case "053"
                            strNewDeptID = "104106" & strEndWageLoc
                        Case "054"
                            strNewDeptID = "104112" & strEndWageLoc
                        Case "055"
                            strNewDeptID = "104116" & strEndWageLoc
                        Case "056"
                            strNewDeptID = "104149" & strEndWageLoc
                        Case "057"
                            strNewDeptID = "105004" & strEndWageLoc
                        Case "058"
                            strNewDeptID = "105011" & strEndWageLoc
                        Case "059"
                            strNewDeptID = "104087" & strEndWageLoc
                        Case "060"
                            strNewDeptID = "105018" & strEndWageLoc
                        Case "061"
                            strNewDeptID = "105034" & strEndWageLoc
                        Case "062"
                            strNewDeptID = "104133" & strEndWageLoc
                        Case "063"
                            strNewDeptID = "105048" & strEndWageLoc
                        Case "064"
                            strNewDeptID = "105049" & strEndWageLoc
                        Case "065"
                            strNewDeptID = "105074" & strEndWageLoc
                        Case "066"
                            strNewDeptID = "105083" & strEndWageLoc
                        Case "067"
                            strNewDeptID = "105087" & strEndWageLoc
                        Case "068"
                            strNewDeptID = "105093" & strEndWageLoc
                        Case "069"
                            strNewDeptID = "114015" & strEndWageLoc
                        Case "070"
                            strNewDeptID = "114025" & strEndWageLoc
                        Case "071"
                            strNewDeptID = "114077" & strEndWageLoc
                        Case "072"
                            strNewDeptID = "102061" & strEndWageLoc
                        Case "073"
                            strNewDeptID = "105064" & strEndWageLoc
                        Case "074"
                            strNewDeptID = "104050" & strEndWageLoc
                        Case "075"
                            strNewDeptID = "104046" & strEndWageLoc
                        Case "076"
                            strNewDeptID = "102090" & strEndWageLoc
                        Case "077"
                            strNewDeptID = "102056" & strEndWageLoc
                        Case "078"
                            strNewDeptID = "102076" & strEndWageLoc
                        Case "079"
                            strNewDeptID = "102012" & strEndWageLoc
                        Case "080"
                            strNewDeptID = "102020" & strEndWageLoc
                        Case "081"
                            strNewDeptID = "102202" & strEndWageLoc
                        Case "082"
                            strNewDeptID = "102204" & strEndWageLoc
                        Case "085"
                            strNewDeptID = "102076" & strEndWageLoc
                        Case "092"
                            strNewDeptID = "102037" & strEndWageLoc
                        Case "093"
                            strNewDeptID = "102039" & strEndWageLoc
                        Case "094"
                            strNewDeptID = "102061" & strEndWageLoc
                        Case "095"
                            strNewDeptID = "102093" & strEndWageLoc
                        Case "096"
                            strNewDeptID = "102098" & strEndWageLoc
                        Case "097"
                            strNewDeptID = "104174" & strEndWageLoc
                        Case "098"
                            strNewDeptID = "104009" & strEndWageLoc
                        Case "100"
                            strNewDeptID = "191015" & strEndWageLoc
                        Case "101"
                            strNewDeptID = "102102" & strEndWageLoc
                        Case "102"
                            strNewDeptID = "102062" & strEndWageLoc
                        Case "103"
                            strNewDeptID = "102073" & strEndWageLoc
                        Case "104"
                            strNewDeptID = "102102" & strEndWageLoc
                        Case "105"
                            strNewDeptID = "102062" & strEndWageLoc
                        Case "106"
                            strNewDeptID = "102062" & strEndWageLoc
                        Case "107"
                            strNewDeptID = "102062" & strEndWageLoc
                        Case "108"
                            strNewDeptID = "102062" & strEndWageLoc
                        Case "109"
                            strNewDeptID = "104031" & strEndWageLoc
                        Case "110"
                            strNewDeptID = "102072" & strEndWageLoc
                        Case "111"
                            strNewDeptID = "104184" & strEndWageLoc
                        Case "112"
                            strNewDeptID = "102040" & strEndWageLoc
                        Case "113"
                            strNewDeptID = "102031" & strEndWageLoc
                        Case "114"
                            strNewDeptID = "102042" & strEndWageLoc
                        Case "115"
                            strNewDeptID = "102457" & strEndWageLoc
                        Case "116"
                            strNewDeptID = "104054" & strEndWageLoc
                        Case "117"
                            strNewDeptID = "102402" & strEndWageLoc
                        Case "300"
                            strNewDeptID = "102102" & strEndWageLoc
                        Case Else
                            strNewDeptID = strFullWageLoc

                    End Select
                Case Else
                    strNewDeptID = strFullWageLoc
            End Select

            Mid(strRecord, 27, 9) = strNewDeptID

            If strSSNIn <> strPreviousSSN Then
                If strEndWageLoc = "099" Or strEndWageLoc = "999" Then bolInHouse = True 'Per Angie E 2/15/13
                If strBegWageLoc = "300" Then bolInHouse = True 'Per Angie E, Angel 2/25/14

                'Get FEIN to check for In-House
                cmdSQL2.CommandText = "select count(*) from Claims cla " &
                    "inner join suta su on cla.intsutaid = su.intsutaid " &
                    "where cla.strssn = '" & Mid(strRecord, 2, 9) & "' and " &
                    "strFEIN in ('13-3568039','13-3726617')"
                cmdSQL2.Connection = objConnection2
                objConnection2.Open()
                x = cmdSQL2.ExecuteScalar()
                objConnection2.Close()

                If x > 0 Then
                    cmdSQL2.CommandText = "select isnull(rtrim(ltrim(max(strPositionType))),'D') from VoltWageEmployeeInfo where strssn = '" & Mid(strRecord, 2, 9) & "'"
                    cmdSQL2.Connection = objConnection2
                    objConnection2.Open()
                    strTemp = cmdSQL2.ExecuteScalar()
                    objConnection2.Close()
                    If strTemp <> "D" Then
                        bolInHouse = True
                    End If
                End If

                If Not bolInHouse Then DailyFile.WriteLine(Replace(strRecord, "–", "/"))
            End If
            strPreviousSSN = strSSNIn
        End While
        objConnection.Close()
        DailyFile.Close()

        'MessageBox.Show("Daily Outgoing file created successfully!")

        lblDailyReport.Text = "Daily Outgoing file created successfully!"
        System.Threading.Thread.Sleep(1000)
        lblDailyReport.Refresh()
        System.Threading.Thread.Sleep(5000)

        Return strFileName
    End Function

    Private Sub UploadReport(path As String)
        Const host = "secureftp6.volt.com"
        Const username = "unempadmin"
        Const password = "C8QacrU4"
        Dim destinationPath = System.IO.Path.GetFileName(path)
        Using session As New Session
            session.Open(New SessionOptions With {
                .Protocol = Protocol.Sftp,
                .GiveUpSecurityAndAcceptAnySshHostKey = True,
                .HostName = host,
                .UserName = username,
                .Password = password})
            Dim transferResults As TransferOperationResult = session.PutFiles(path, destinationPath)
            transferResults.Check()
        End Using
    End Sub

    Public Sub Archive(path As String)
        Const archvieFolder = "C:\VoltTransmittedFiles"
        Dim archviePath = IO.Path.Combine(archvieFolder, IO.Path.GetFileName(path))
        File.Move(path, archviePath)
    End Sub


End Class
