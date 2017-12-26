Imports System.Data.SqlClient


Public Class Form1
    Dim ActiveStation = 0
    Dim user = Environment.UserName
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\" + user + "\Documents\OPF-TableNo.txt")
    Public TransferCase
    Dim s2case = 0
    Dim s3case = 0
    Dim s4case = 0
    Dim s5case = 0
    Dim s6case = 0
    Dim s7case = 0
    Dim s8case = 0
    Dim s9case = 0
    Dim s10case = 0
    Dim s11case = 0

    Dim EndShiftStatus = 0

    Dim con2 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

    Dim con3 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

    Private Sub ChangeEmployee(e, f, g)
        '''''''VARIABLES'''''''
        Dim con As SqlConnection

        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMHumanResource;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim cEmpquery As String = "SELECT EN.EmployeeNumber as Username, CONCAT( E.[LastName], ', ', E.[FirstName], ' - ', EN.EmployeeNumber ) as Fullname FROM [SMHumanResource].[dbo].[Employee] as E LEFT JOIN EmployeeNumber as EN On EN.EmpID = E.EmpID WHERE E.[LastName] <> '' AND EN.EmployeeNumber = @emp ORDER BY FullName"
        Dim cEmpcmd As SqlCommand = New SqlCommand(cEmpquery, con)
        cEmpcmd.Parameters.AddWithValue("@emp", e.SelectedValue.ToString)
        con.Open()
        Using reader As SqlDataReader = cEmpcmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()

                End While

            End If
        End Using
        con.Close()
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            If (ComboBox14.Text <> "") Then
                ComboBox14.Visible = False
                Label21.Visible = True
                Label21.Text = ComboBox14.Text
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GetLeadman()
        GetTableDetails()
        If (Button13.Text = "Release") Then
            Timer3.Enabled = True
        End If

    End Sub

    Private Sub Clock_Tick(sender As Object, e As EventArgs) Handles Clock.Tick
        Label1.Text = Format(Now, "hh:mm:ss")
        Label2.Text = Format(Now, "MMMM dd, yyyy")

        Dim TableSetID = 0
        '''''''QUERY FOR SELECTING ACTIVE TABLE'''''''''''
        Dim TableSetIDquery As String = "SELECT TableSetID FROM [SMProduction].[dbo].[TableSet] WHERE TableID = @TS AND TableSetStatus = 1"
        Dim TableSetIDquerycmd As SqlCommand = New SqlCommand(TableSetIDquery, con3)
        TableSetIDquerycmd.Parameters.AddWithValue("@TS", TableSet)
        con3.Open()
        Using reader As SqlDataReader = TableSetIDquerycmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    TableSetID = reader.Item("TableSetID").ToString
                End While

            End If
        End Using
        con3.Close()



        '''''''CHECK TABLE GOAL'''''''''''
        Dim TableGoal As String = "Select (DateDiff(Minute, [TableSetTimeIn], GETDATE()) / 20) As goal From [SMProduction].[dbo].[TableSet] as TS Where TS.TableID = @TS And TS.TableSetStatus = 1 AND TS.TableSetID = @TSID"
        Dim TableGoalQuery As SqlCommand = New SqlCommand(TableGoal, con3)
        TableGoalQuery.Parameters.AddWithValue("@TS", TableSet)
        TableGoalQuery.Parameters.AddWithValue("@TSID", TableSetID)
        con3.Open()
        Using reader As SqlDataReader = TableGoalQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    Label79.Text = reader.Item("goal")
                End While

            End If
        End Using
        con3.Close()

        '''''''CHECK TABLE DONE'''''''''''
        Dim TableAchieved As String = "SELECT COUNT(SomtrackID) as achieved FROM [SMProduction].[dbo].[ProductionHead] WHERE TableNo = @TS AND StationID = 0 AND TableSetID = @TSID"
        Dim TableAchievedQuery As SqlCommand = New SqlCommand(TableAchieved, con3)
        TableAchievedQuery.Parameters.AddWithValue("@TS", TableSet)
        TableAchievedQuery.Parameters.AddWithValue("@TSID", TableSetID)
        con3.Open()
        Using reader As SqlDataReader = TableAchievedQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    Label77.Text = reader.Item("achieved")
                End While

            End If
        End Using
        con3.Close()
        GetPending()
        GetActive()
    End Sub
    Private Sub GetActive()



        '''''''CHECK TABLE CASES'''''''''''
        Dim TableActiveCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PD.Status = 1 And PH.StationID Is Not NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableActiveCaseQuery As SqlCommand = New SqlCommand(TableActiveCase, con3)
        TableActiveCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = TableActiveCaseQuery.ExecuteReader()
            If reader.HasRows Then
                EndShiftStatus = 0

                Dim s1Active = 0
                Dim s2Active = 0
                Dim s3Active = 0
                Dim s4Active = 0
                Dim s5Active = 0
                Dim s6Active = 0
                Dim s7Active = 0
                Dim s8Active = 0
                Dim s9Active = 0
                Dim s10Active = 0
                Dim s11Active = 0
                Dim s12Active = 0

                While reader.Read()

                    If (reader.Item("StationID") = 1) Then
                        Label18.Text = reader.Item("SomtrackID").ToString
                        Button29.Visible = True
                        s1Active = 1
                    ElseIf (reader.Item("StationID") = 2) Then
                        Label34.Text = reader.Item("SomtrackID").ToString
                        Button30.Visible = True
                        s2Active = 1
                    ElseIf (reader.Item("StationID") = 3) Then
                        Label39.Text = reader.Item("SomtrackID").ToString
                        Button31.Visible = True
                        s3Active = 1
                    ElseIf (reader.Item("StationID") = 4) Then
                        Label22.Text = reader.Item("SomtrackID").ToString
                        Button32.Visible = True
                        s4Active = 1
                    ElseIf (reader.Item("StationID") = 5) Then
                        Label44.Text = reader.Item("SomtrackID").ToString
                        Button33.Visible = True
                        s5Active = 1
                    ElseIf (reader.Item("StationID") = 6) Then
                        Label29.Text = reader.Item("SomtrackID").ToString
                        Button34.Visible = True
                        s6Active = 1
                    ElseIf (reader.Item("StationID") = 7) Then
                        Label74.Text = reader.Item("SomtrackID").ToString
                        Button35.Visible = True
                        s7Active = 1
                    ElseIf (reader.Item("StationID") = 8) Then
                        Label64.Text = reader.Item("SomtrackID").ToString
                        Button36.Visible = True
                        s8Active = 1
                    ElseIf (reader.Item("StationID") = 9) Then
                        Label69.Text = reader.Item("SomtrackID").ToString
                        Button37.Visible = True
                        s9Active = 1
                    ElseIf (reader.Item("StationID") = 10) Then
                        Label54.Text = reader.Item("SomtrackID").ToString
                        Button38.Visible = True
                        s10Active = 1
                    ElseIf (reader.Item("StationID") = 11) Then
                        Label59.Text = reader.Item("SomtrackID").ToString
                        Button39.Visible = True
                        s11Active = 1
                    ElseIf (reader.Item("StationID") = 12) Then
                        Label49.Text = reader.Item("SomtrackID").ToString
                        Button40.Visible = True
                        s12Active = 1
                    End If


                End While

                If s1Active = 0 Then
                    Button29.Visible = False
                    Label18.Text = ""
                End If
                If s2Active = 0 Then
                    Button30.Visible = False
                    Label34.Text = ""
                End If
                If s3Active = 0 Then
                    Button31.Visible = False
                    Label39.Text = ""
                End If
                If s4Active = 0 Then
                    Button32.Visible = False
                    Label22.Text = ""
                End If
                If s5Active = 0 Then
                    Button33.Visible = False
                    Label44.Text = ""
                End If
                If s6Active = 0 Then
                    Button34.Visible = False
                    Label29.Text = ""
                End If
                If s7Active = 0 Then
                    Button35.Visible = False
                    Label74.Text = ""
                End If
                If s8Active = 0 Then
                    Button36.Visible = False
                    Label64.Text = ""
                End If
                If s9Active = 0 Then
                    Button37.Visible = False
                    Label69.Text = ""
                End If
                If s10Active = 0 Then
                    Button38.Visible = False
                    Label54.Text = ""
                End If
                If s11Active = 0 Then
                    Button39.Visible = False
                    Label59.Text = ""
                End If
                If s12Active = 0 Then
                    Button40.Visible = False
                    Label49.Text = ""
                End If





            Else
                EndShiftStatus = 1
            End If
        End Using
        con3.Close()



    End Sub
    Private Sub GetPending()


        s2case = 0
        s3case = 0
        s4case = 0
        s5case = 0
        s6case = 0
        s7case = 0
        s8case = 0
        s9case = 0
        s10case = 0
        s11case = 0


        '''''''CHECK TABLE CASES'''''''''''
        Dim TableCase As String = "SELECT PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PD.Status = 2 AND PH.StationID IS NOT NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableCaseQuery As SqlCommand = New SqlCommand(TableCase, con3)
        TableCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = TableCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 2) Then
                        s2case = s2case + 1
                    ElseIf (reader.Item("StationID") = 3) Then
                        s3case = s3case + 1

                    ElseIf (reader.Item("StationID") = 4) Then
                        s4case = s4case + 1

                    ElseIf (reader.Item("StationID") = 5) Then
                        s5case = s5case + 1

                    ElseIf (reader.Item("StationID") = 6) Then
                        s6case = s6case + 1

                    ElseIf (reader.Item("StationID") = 7) Then
                        s7case = s7case + 1

                    ElseIf (reader.Item("StationID") = 8) Then
                        s8case = s8case + 1

                    ElseIf (reader.Item("StationID") = 9) Then
                        s9case = s9case + 1

                    ElseIf (reader.Item("StationID") = 10) Then
                        s10case = s10case + 1

                    ElseIf (reader.Item("StationID") = 11) Then
                        s11case = s11case + 1


                    End If
                End While

            End If
        End Using

        Label32.Text = s2case
        Label37.Text = s3case
        Label24.Text = s4case
        Label42.Text = s5case
        Label27.Text = s6case
        Label72.Text = s7case
        Label62.Text = s8case
        Label67.Text = s9case
        Label52.Text = s10case
        Label57.Text = s11case
        Label47.Text = s11case

        con3.Close()
    End Sub

    Private Sub GetTableDetails()
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim con5 As SqlConnection
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        con5 = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMHumanResource;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING ROSTER'''''''''''
        Dim empquery As String = "SELECT EN.EmployeeNumber as EmployeeID, CONCAT( [LastName], ', ', [FirstName], ' - ', EN.EmployeeNumber ) as FullName FROM [SMProduction].[dbo].[TableMembers] as TM LEFT JOIN [SMProduction].[dbo].[TableSet] as TS ON TS.TableSetID = TM.TableSetID LEFT JOIN [SMHumanResource].[dbo].[EmployeeNumber] as EN ON EN.EmployeeNumber = TM.EmployeeID LEFT JOIN [SMHumanResource].[dbo].[Employee] as E ON E.EmpID = EN.EmpID WHERE TS.TableID = @TN AND TableMemberStatus = 1 AND StationID = @SN AND TS.TableSetStatus = 1"
        Dim empcmd As SqlCommand = New SqlCommand(empquery, con)


        Dim empquery2 As String = "SELECT '0' as Username, '' as Fullname UNION SELECT EN.EmployeeNumber as Username, CONCAT( E.[LastName], ', ', E.[FirstName], ' - ', EN.EmployeeNumber ) as Fullname FROM [SMHumanResource].[dbo].[Employee] as E LEFT JOIN EmployeeNumber as EN On EN.EmpID = E.EmpID WHERE E.[LastName] <> '' ORDER BY FullName"
        Dim empcmd2 As SqlCommand = New SqlCommand(empquery2, con5)


        For i As Integer = 2 To 13

            empcmd.Parameters.Clear()
            empcmd.Parameters.AddWithValue("@TN", TableSet)
            empcmd.Parameters.AddWithValue("@SN", i - 1)
            con.Close()
            con.Open()
            Using reader As SqlDataReader = empcmd.ExecuteReader()

                If reader.HasRows Then
                    Dim dt As DataTable = New DataTable
                    dt.Load(reader)

                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "EmployeeID"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("Button" + (i + 12).ToString()), Button).Text = "Sign Out"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("Button" + (i + 12).ToString()), Button).BackColor = Color.RosyBrown
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).Enabled = False
                    CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).BackColor = Color.DarkSlateGray
                    ActiveStation = ActiveStation + 1
                Else
                    con.Close()
                    con5.Open()

                    Using reader2 As SqlDataReader = empcmd2.ExecuteReader()
                        Dim dt2 As DataTable = New DataTable
                        dt2.Load(reader2)

                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt2
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "Username"
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"

                    End Using
                    con5.Close()
                End If
            End Using
        Next

        con.Close()

    End Sub
    Private Sub GetLeadman()
        '''''''VARIABLES'''''''
        Dim con As SqlConnection

        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim leadquery As String = "SELECT '0' as Username, '' as FullName UNION SELECT [Username] ,CONCAT([LastName],', ',[FirstName], ' - ', [Username]) as FullName FROM [Users] WHERE DepID = '9' AND Status = 'Active' ORDER BY FullName"
        Dim leadcmd As SqlCommand = New SqlCommand(leadquery, con)

        con.Open()

        Using reader As SqlDataReader = leadcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)
            ComboBox1.DataSource = dt
            ComboBox1.ValueMember = "Username"
            ComboBox1.DisplayMember = "FullName"

        End Using
        con.Close()



        con.Open()
        Dim leadcurrquery As String = "SELECT TOP 1 [TableSetLeadID] FROM [SMProduction].[dbo].[TableSet] WHERE TableID = @TN AND TableSetStatus = 1"
        Dim leadcurrcmd As SqlCommand = New SqlCommand(leadcurrquery, con)
        leadcurrcmd.Parameters.AddWithValue("@TN", TableSet)

        Using reader As SqlDataReader = leadcurrcmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    ComboBox1.SelectedValue = reader.Item("TableSetLeadID")
                    Button13.BackColor = Color.DarkGray
                    Button13.Text = "Release"
                    ComboBox1.Enabled = False
                    Button14.Enabled = False
                    Button15.Enabled = False
                    Button16.Enabled = False
                    Button20.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                    Button19.Enabled = False
                    Button21.Enabled = False
                    Button22.Enabled = False
                    Button23.Enabled = False
                    Button24.Enabled = False
                    Button25.Enabled = False
                End While
            End If

        End Using
        con.Close()




    End Sub
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If s2case > 3 Then
            If Panel2.BackColor = Color.DarkSlateGray Then
                Panel2.BackColor = Color.Red
            Else
                Panel2.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel2.BackColor = Color.DarkSlateGray
        End If
        If s3case > 3 Then
            If Panel3.BackColor = Color.DarkSlateGray Then
                Panel3.BackColor = Color.Red
            Else
                Panel3.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel3.BackColor = Color.DarkSlateGray
        End If
        If s4case > 3 Then
            If Panel4.BackColor = Color.DarkSlateGray Then
                Panel4.BackColor = Color.Red
            Else
                Panel4.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel4.BackColor = Color.DarkSlateGray
        End If
        If s5case > 3 Then
            If Panel5.BackColor = Color.DarkSlateGray Then
                Panel5.BackColor = Color.Red
            Else
                Panel5.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel5.BackColor = Color.DarkSlateGray
        End If
        If s6case > 3 Then
            If Panel6.BackColor = Color.DarkSlateGray Then
                Panel6.BackColor = Color.Red
            Else
                Panel6.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel6.BackColor = Color.DarkSlateGray
        End If
        If s7case > 3 Then
            If Panel7.BackColor = Color.DarkSlateGray Then
                Panel7.BackColor = Color.Red
            Else
                Panel7.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel7.BackColor = Color.DarkSlateGray
        End If
        If s8case > 3 Then
            If Panel8.BackColor = Color.DarkSlateGray Then
                Panel8.BackColor = Color.Red
            Else
                Panel8.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel8.BackColor = Color.DarkSlateGray
        End If
        If s9case > 3 Then
            If Panel9.BackColor = Color.DarkSlateGray Then
                Panel9.BackColor = Color.Red
            Else
                Panel9.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel9.BackColor = Color.DarkSlateGray
        End If
        If s10case > 3 Then
            If Panel10.BackColor = Color.DarkSlateGray Then
                Panel10.BackColor = Color.Red
            Else
                Panel10.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel10.BackColor = Color.DarkSlateGray
        End If
        If s11case > 3 Then
            If Panel11.BackColor = Color.DarkSlateGray Then
                Panel11.BackColor = Color.Red
                Panel12.BackColor = Color.Red
            Else
                Panel11.BackColor = Color.DarkSlateGray
                Panel12.BackColor = Color.DarkSlateGray
            End If
        Else
            Panel11.BackColor = Color.DarkSlateGray
            Panel12.BackColor = Color.DarkSlateGray
        End If
    End Sub

    Private Sub BreakStatus(x)
        If (x.Text = "Break Time") Then
            x.Text = "On Break"
            x.BackColor = Color.Orange
        Else
            x.Text = "Break Time"
            x.BackColor = Color.LawnGreen
        End If

    End Sub

    Private Sub SignStatus(x, e, f)
        If (x.Text = "Sign In") Then
            If (e.Text = "") Then
                MessageBox.Show("Employee Name cannot be empty", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            Else

                '''''''VARIABLES'''''''
                Dim con As SqlConnection
                Dim Station = f.Name.Substring(5, 1)
                Dim SelectedEmp = e.SelectedValue
                Dim readTableSetID = ""


                'SERVER'
                con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

                Dim CurrentHead As String = "SELECT TOP (1) TableSetID FROM TableSet WHERE TableID = @TS AND TableSetStatus = 1 ORDER BY TableSetID DESC"
                Dim CurrentHeadQuery As SqlCommand = New SqlCommand(CurrentHead, con)
                CurrentHeadQuery.Parameters.AddWithValue("@TS", TableSet)

                con.Open()
                Using reader As SqlDataReader = CurrentHeadQuery.ExecuteReader()
                    If reader.HasRows Then

                        While reader.Read()
                            readTableSetID = reader.Item("TableSetID")
                        End While

                        con.Close()
                        con.Open()
                        Dim TableDetails As String = "UPDATE TM SET TableMemberTimeOut = GETDATE(), TableMemberStatus = '3' FROM TableMembers as TM LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 AND TS.TableID = @TS AND TM.StationID = @SID"
                        Dim TableDetailsQuery As SqlCommand = New SqlCommand(TableDetails, con)
                        TableDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
                        TableDetailsQuery.Parameters.AddWithValue("@SID", Station)
                        TableDetailsQuery.ExecuteNonQuery()
                        con.Close()

                        con.Open()
                        Dim TableAddDetails As String = "INSERT INTO TableMembers (TableSetID,EmployeeID, StationID, TableMemberTimeIn, TableMemberStatus) Values (@TSID, @Emp, @SID,GETDATE(), '1')"
                        Dim TableAddDetailsQuery As SqlCommand = New SqlCommand(TableAddDetails, con)
                        TableAddDetailsQuery.Parameters.AddWithValue("@Emp", SelectedEmp)
                        TableAddDetailsQuery.Parameters.AddWithValue("@TSID", readTableSetID)
                        TableAddDetailsQuery.Parameters.AddWithValue("@SID", Station)
                        TableAddDetailsQuery.ExecuteNonQuery()
                        con.Close()


                    End If

                End Using


                x.Text = "Sign Out"
                x.BackColor = Color.RosyBrown
                e.Enabled = False
                f.BackColor = Color.DarkSlateGray
                ActiveStation = ActiveStation + 1
            End If
        Else

            '''''''VARIABLES'''''''
            Dim con As SqlConnection

            'SERVER'
            con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMHumanResource;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

            '''''''QUERY FOR SELECTING LEADMAN'''''''''''
            Dim empquery As String = "SELECT '0' as Username, '' as Fullname UNION SELECT EN.EmployeeNumber as Username, CONCAT( E.[LastName], ', ', E.[FirstName], ' - ', EN.EmployeeNumber ) as Fullname FROM [SMHumanResource].[dbo].[Employee] as E LEFT JOIN EmployeeNumber as EN On EN.EmpID = E.EmpID WHERE E.[LastName] <> '' ORDER BY FullName"
            Dim empcmd As SqlCommand = New SqlCommand(empquery, con)
            Dim empValue = e.SelectedValue
            con.Open()
            Using reader As SqlDataReader = empcmd.ExecuteReader()
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                e.DataSource = dt
                e.ValueMember = "Username"
                e.DisplayMember = "FullName"


            End Using
            con.Close()

            e.selectedvalue = empValue
            x.Text = "Sign In"
            x.BackColor = Color.LightGray
            e.Enabled = True
            f.BackColor = Color.DarkGray
            ActiveStation = ActiveStation - 1
        End If

    End Sub

    Private Sub GetEmployee()
        '''''''VARIABLES'''''''
        Dim con As SqlConnection

        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMHumanResource;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim empquery As String = "SELECT '0' as Username, '' as Fullname UNION SELECT EN.EmployeeNumber as Username, CONCAT( E.[LastName], ', ', E.[FirstName], ' - ', EN.EmployeeNumber ) as Fullname FROM [SMHumanResource].[dbo].[Employee] as E LEFT JOIN EmployeeNumber as EN On EN.EmpID = E.EmpID WHERE E.[LastName] <> '' ORDER BY FullName"
        Dim empcmd As SqlCommand = New SqlCommand(empquery, con)

        con.Open()

        For i As Integer = 2 To 13


            Using reader As SqlDataReader = empcmd.ExecuteReader()
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt
                CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "Username"
                CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"

            End Using
        Next

        con.Close()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        For i As Integer = 1 To 12
            CType(CType(Me.Controls("Panel" + (i).ToString()), Panel).Controls("Button" + (i).ToString()), Button).Text = "On Break"
            CType(CType(Me.Controls("Panel" + (i).ToString()), Panel).Controls("Button" + (i).ToString()), Button).BackColor = Color.Orange

        Next



    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        For i As Integer = 1 To 12
            CType(CType(Me.Controls("Panel" + (i).ToString()), Panel).Controls("Button" + (i).ToString()), Button).Text = "Break Time"
            CType(CType(Me.Controls("Panel" + (i).ToString()), Panel).Controls("Button" + (i).ToString()), Button).BackColor = Color.LawnGreen


        Next

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        BreakStatus(Button11)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        BreakStatus(Button7)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        BreakStatus(Button1)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        BreakStatus(Button12)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        BreakStatus(Button2)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        BreakStatus(Button10)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        BreakStatus(Button3)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        BreakStatus(Button9)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        BreakStatus(Button4)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        BreakStatus(Button8)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        BreakStatus(Button5)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        BreakStatus(Button6)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        SignStatus(Button14, ComboBox2, Panel1)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        SignStatus(Button17, ComboBox5, Panel4)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        SignStatus(Button15, ComboBox3, Panel2)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        SignStatus(Button18, ComboBox6, Panel5)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        SignStatus(Button19, ComboBox7, Panel6)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        SignStatus(Button16, ComboBox4, Panel3)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        SignStatus(Button20, ComboBox8, Panel7)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        SignStatus(Button21, ComboBox9, Panel8)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        SignStatus(Button22, ComboBox10, Panel9)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        SignStatus(Button23, ComboBox11, Panel10)
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        SignStatus(Button25, ComboBox13, Panel12)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        SignStatus(Button24, ComboBox12, Panel11)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If (Button13.Text = "Lock In") Then
            If (ComboBox1.Text = "") Then
                MessageBox.Show("Leadman cannot be empty", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            Else
                Timer3.Enabled = True

                If MessageBox.Show("Are you sure you want to lock in this roster?", "Title", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                    If (ActiveStation = 12) Then

                        If (ComboBox1.Enabled = True) Then


                            '''''''VARIABLES'''''''
                            Dim con As SqlConnection
                            'SERVER'
                            con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
                            con.Open()


                            '''''''CHECK FOR ACTIVE TABLE''''''
                            Dim GetActiveTable As String = "SELECT TOP (1) TableSetID FROM TableSet WHERE TableID = @TID AND TableSetStatus = 1 ORDER BY TableSetID DESC"
                            Dim GetActiveTableQuery As SqlCommand = New SqlCommand(GetActiveTable, con)
                            GetActiveTableQuery.Parameters.AddWithValue("@TID", TableSet)
                            Dim ActiveTable = ""

                            Using reader As SqlDataReader = GetActiveTableQuery.ExecuteReader()
                                If reader.HasRows Then
                                    While reader.Read()
                                        ActiveTable = reader.Item("TableSetID")
                                    End While
                                End If

                            End Using


                            If ActiveTable = "" Then

                                Dim TableHead As String = "INSERT INTO TableSet (TableID,TableSetLeadID, TableSetName, TableSetTimeIn, TableSetStatus) Values (@TS, @LM, @SN,GETDATE(), '1')"
                                Dim TableHeadQuery As SqlCommand = New SqlCommand(TableHead, con)
                                TableHeadQuery.Parameters.AddWithValue("@LM", ComboBox1.SelectedValue)
                                TableHeadQuery.Parameters.AddWithValue("@TS", TableSet)
                                TableHeadQuery.Parameters.AddWithValue("@SN", Label21.Text)
                                TableHeadQuery.ExecuteNonQuery()


                                Dim TableMembers As String = "INSERT INTO TableMembers (TableSetID,EmployeeID, StationID, TableMemberTimeIn, TableMemberStatus) Values (@TSID, @Emp1, '1',GETDATE(), '1'), (@TSID, @Emp2, '2',GETDATE(), '1'), (@TSID, @Emp3, '3',GETDATE(), '1'), (@TSID, @Emp4, '4',GETDATE(), '1'), (@TSID, @Emp5, '5',GETDATE(), '1'), (@TSID, @Emp6, '6',GETDATE(), '1'), (@TSID, @Emp7, '7',GETDATE(), '1'), (@TSID, @Emp8, '8',GETDATE(), '1'), (@TSID, @Emp9, '9',GETDATE(), '1'), (@TSID, @Emp10, '10',GETDATE(), '1'), (@TSID, @Emp11, '11',GETDATE(), '1'), (@TSID, @Emp12, '12',GETDATE(), '1')"
                                Dim TableMembersQuery As SqlCommand = New SqlCommand(TableMembers, con)


                                Dim CurrentHead As String = "SELECT TOP (1) TableSetID FROM TableSet WHERE TableID = @TID ORDER BY TableSetID DESC"
                                Dim CurrentHeadQuery As SqlCommand = New SqlCommand(CurrentHead, con)
                                CurrentHeadQuery.Parameters.AddWithValue("@TID", TableSet)
                                Dim readTableSetID = ""

                                Using reader As SqlDataReader = CurrentHeadQuery.ExecuteReader()
                                    If reader.HasRows Then
                                        While reader.Read()
                                            readTableSetID = reader.Item("TableSetID")
                                        End While
                                    End If

                                End Using

                                TableMembersQuery.Parameters.AddWithValue("@TSID", readTableSetID)

                                TableMembersQuery.Parameters.AddWithValue("@Emp1", ComboBox2.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp2", ComboBox3.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp3", ComboBox4.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp4", ComboBox5.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp5", ComboBox6.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp6", ComboBox7.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp7", ComboBox8.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp8", ComboBox9.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp9", ComboBox10.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp10", ComboBox11.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp11", ComboBox12.SelectedValue)
                                TableMembersQuery.Parameters.AddWithValue("@Emp12", ComboBox13.SelectedValue)
                                TableMembersQuery.ExecuteNonQuery()
                                con.Close()


                                Button13.BackColor = Color.DarkGray
                                Button13.Text = "Release"
                                ComboBox1.Enabled = False
                                Button14.Enabled = False
                                Button15.Enabled = False
                                Button16.Enabled = False
                                Button20.Enabled = False
                                Button17.Enabled = False
                                Button18.Enabled = False
                                Button19.Enabled = False
                                Button21.Enabled = False
                                Button22.Enabled = False
                                Button23.Enabled = False
                                Button24.Enabled = False
                                Button25.Enabled = False
                            Else
                                Timer3.Enabled = False
                                Button13.BackColor = Color.LightGray
                                Button13.Text = "Lock In"
                                'ComboBox1.Enabled = True
                                Button14.Enabled = True
                                Button15.Enabled = True
                                Button16.Enabled = True
                                Button20.Enabled = True
                                Button17.Enabled = True
                                Button18.Enabled = True
                                Button19.Enabled = True
                                Button21.Enabled = True
                                Button22.Enabled = True
                                Button23.Enabled = True
                                Button24.Enabled = True
                                Button25.Enabled = True
                                Button28.Enabled = True

                                Panel2.BackColor = Color.DarkSlateGray
                                Panel3.BackColor = Color.DarkSlateGray
                                Panel4.BackColor = Color.DarkSlateGray
                                Panel5.BackColor = Color.DarkSlateGray
                                Panel6.BackColor = Color.DarkSlateGray
                                Panel7.BackColor = Color.DarkSlateGray
                                Panel8.BackColor = Color.DarkSlateGray
                                Panel9.BackColor = Color.DarkSlateGray
                                Panel10.BackColor = Color.DarkSlateGray
                                Panel11.BackColor = Color.DarkSlateGray
                                Panel12.BackColor = Color.DarkSlateGray
                            End If
                        Else
                            Button13.BackColor = Color.DarkGray
                            Button13.Text = "Release"
                            ComboBox1.Enabled = False
                            Button14.Enabled = False
                            Button15.Enabled = False
                            Button16.Enabled = False
                            Button20.Enabled = False
                            Button17.Enabled = False
                            Button18.Enabled = False
                            Button19.Enabled = False
                            Button21.Enabled = False
                            Button22.Enabled = False
                            Button23.Enabled = False
                            Button24.Enabled = False
                            Button25.Enabled = False
                            Button28.Enabled = False
                        End If
                    Else
                        MessageBox.Show("Please fill all the station")
                    End If
                End If
            End If

        Else
            Timer3.Enabled = False
            Button13.BackColor = Color.LightGray
            Button13.Text = "Lock In"
            'ComboBox1.Enabled = True
            Button14.Enabled = True
            Button15.Enabled = True
            Button16.Enabled = True
            Button20.Enabled = True
            Button17.Enabled = True
            Button18.Enabled = True
            Button19.Enabled = True
            Button21.Enabled = True
            Button22.Enabled = True
            Button23.Enabled = True
            Button24.Enabled = True
            Button25.Enabled = True
            Button28.Enabled = True

            Panel2.BackColor = Color.DarkSlateGray
            Panel3.BackColor = Color.DarkSlateGray
            Panel4.BackColor = Color.DarkSlateGray
            Panel5.BackColor = Color.DarkSlateGray
            Panel6.BackColor = Color.DarkSlateGray
            Panel7.BackColor = Color.DarkSlateGray
            Panel8.BackColor = Color.DarkSlateGray
            Panel9.BackColor = Color.DarkSlateGray
            Panel10.BackColor = Color.DarkSlateGray
            Panel11.BackColor = Color.DarkSlateGray
            Panel12.BackColor = Color.DarkSlateGray

        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ChangeEmployee(ComboBox2, Label20, Label18)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ChangeEmployee(ComboBox3, Label34, Label32)
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ChangeEmployee(ComboBox4, Label39, Label37)
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ChangeEmployee(ComboBox5, Label24, Label22)
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ChangeEmployee(ComboBox6, Label44, Label42)
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ChangeEmployee(ComboBox7, Label29, Label27)
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        ChangeEmployee(ComboBox8, Label74, Label72)
    End Sub
    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        ChangeEmployee(ComboBox9, Label64, Label62)
    End Sub
    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        ChangeEmployee(ComboBox10, Label69, Label67)
    End Sub
    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        ChangeEmployee(ComboBox11, Label54, Label52)
    End Sub
    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        ChangeEmployee(ComboBox12, Label59, Label57)
    End Sub
    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        ChangeEmployee(ComboBox13, Label49, Label47)
    End Sub
    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        con.Open()
        '''''''CHECK FOR ACTIVE TABLE''''''
        Dim GetActiveTable As String = "SELECT Distinct [TableSetName], MAX(TableSetID) as TableSetID FROM TableSet WHERE TableID = @TID AND TableSetStatus = 3 GROUP BY [TableSetName] ORDER BY [TableSetName] asc"
        Dim GetActiveTableQuery As SqlCommand = New SqlCommand(GetActiveTable, con)
        GetActiveTableQuery.Parameters.AddWithValue("@TID", TableSet)
        Dim ActiveTable = ""

        Using reader As SqlDataReader = GetActiveTableQuery.ExecuteReader()
            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox14.DataSource = dt
            ComboBox14.ValueMember = "TableSetID"
            ComboBox14.DisplayMember = "TableSetName"


        End Using
        con.Close()

        Dim con5 As SqlConnection
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        con5 = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMHumanResource;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING ROSTER'''''''''''
        Dim empquery As String = "SELECT EN.EmployeeNumber as EmployeeID, CONCAT( [LastName], ', ', [FirstName], ' - ', EN.EmployeeNumber ) as FullName FROM [SMProduction].[dbo].[TableMembers] as TM LEFT JOIN [SMProduction].[dbo].[TableSet] as TS ON TS.TableSetID = TM.TableSetID LEFT JOIN [SMHumanResource].[dbo].[EmployeeNumber] as EN ON EN.EmployeeNumber = TM.EmployeeID LEFT JOIN [SMHumanResource].[dbo].[Employee] as E ON E.EmpID = EN.EmpID WHERE TS.TableSetID = @TSID AND StationID = @SN"
        Dim empcmd As SqlCommand = New SqlCommand(empquery, con)
        empcmd.Parameters.AddWithValue("@TSID", ComboBox14.SelectedValue)


        Dim empquery2 As String = "SELECT '0' as Username, '' as Fullname UNION SELECT EN.EmployeeNumber as Username, CONCAT( E.[LastName], ', ', E.[FirstName], ' - ', EN.EmployeeNumber ) as Fullname FROM [SMHumanResource].[dbo].[Employee] as E LEFT JOIN EmployeeNumber as EN On EN.EmpID = E.EmpID WHERE E.[LastName] <> '' ORDER BY FullName"
        Dim empcmd2 As SqlCommand = New SqlCommand(empquery2, con5)


        For i As Integer = 2 To 13

            empcmd.Parameters.Clear()
            empcmd.Parameters.AddWithValue("@TN", TableSet)
            empcmd.Parameters.AddWithValue("@SN", i - 1)
            con.Close()
            con.Open()
            Using reader As SqlDataReader = empcmd.ExecuteReader()

                If reader.HasRows Then
                    Dim dt As DataTable = New DataTable
                    dt.Load(reader)

                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "EmployeeID"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("Button" + (i + 12).ToString()), Button).Text = "Sign Out"
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("Button" + (i + 12).ToString()), Button).BackColor = Color.RosyBrown
                    CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).Enabled = False
                    CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).BackColor = Color.DarkSlateGray
                    ActiveStation = ActiveStation + 1
                Else
                    con.Close()
                    con5.Open()

                    Using reader2 As SqlDataReader = empcmd2.ExecuteReader()
                        Dim dt2 As DataTable = New DataTable
                        dt2.Load(reader2)

                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt2
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "Username"
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"

                    End Using
                    con5.Close()
                End If
            End Using
        Next

        con.Close()





































        ComboBox14.Visible = True
        Label21.Visible = False
        ComboBox14.Text = Label21.Text
        ComboBox14.Select()
    End Sub



    Private Sub Label79_Click(sender As Object, e As EventArgs) Handles Label79.Click

    End Sub

    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click

    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        If Button13.Text = "Lock In" Then
            If EndShiftStatus = 1 Then
                If MessageBox.Show("End shift will stop the One Piece Flow. Please make sure that all case are processed or in queue.", "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then

                    '''''''VARIABLES'''''''
                    Dim con As SqlConnection
                    'SERVER'
                    con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

                    con.Open()
                    Dim TableDetails As String = "UPDATE TM SET TableMemberTimeOut = GETDATE(), TableMemberStatus = '3' FROM TableMembers as TM LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 AND TS.TableID = @TS"
                    Dim TableDetailsQuery As SqlCommand = New SqlCommand(TableDetails, con)
                    TableDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
                    TableDetailsQuery.ExecuteNonQuery()
                    con.Close()

                    con.Open()
                    Dim TableHead As String = "UPDATE TableSet SET TableSetTimeOut = GETDATE(), TableSetStatus = '3' WHERE TableSetStatus = 1 AND TableID = @TS"
                    Dim TableHeadQuery As SqlCommand = New SqlCommand(TableHead, con)
                    TableHeadQuery.Parameters.AddWithValue("@TS", TableSet)
                    TableHeadQuery.ExecuteNonQuery()
                    con.Close()
                    Me.Controls.Clear()
                    Me.InitializeComponent()
                    GetLeadman()
                    GetTableDetails()
                    Application.Restart()
                    ComboBox1.Enabled = True

                End If
            Else
                MessageBox.Show("There are still active cases on this shift")
            End If
        Else
                MessageBox.Show("Release Roster first")
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged

    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click

    End Sub

    Private Sub Label54_Click(sender As Object, e As EventArgs) Handles Label54.Click

    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TransferCase = Label18.Text
        Form2.Station = "1"
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click

        TransferCase = Label34.Text
        Form2.Station = "2"
        Form2.Show()
        Form2.Select()

    End Sub
    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Form2.Station = "3"
        TransferCase = Label39.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        Form2.Station = "4"
        TransferCase = Label22.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Form2.Station = "5"
        TransferCase = Label44.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Form2.Station = "6"
        TransferCase = Label29.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Form2.Station = "7"
        TransferCase = Label74.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Form2.Station = "8"
        TransferCase = Label64.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Form2.Station = "9"
        TransferCase = Label69.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        Form2.Station = "10"
        TransferCase = Label54.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Form2.Station = "11"
        TransferCase = Label59.Text
        Form2.Show()
        Form2.Select()
    End Sub
    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Form2.Station = "12"
        TransferCase = Label49.Text
        Form2.Show()
        Form2.Select()
    End Sub

    Private Sub Form1_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick

        If Application.OpenForms().OfType(Of Form2).Any Then
            Form2.Select()
        End If
    End Sub

    Private Sub Label57_Click(sender As Object, e As EventArgs) Handles Label57.Click

    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox14.SelectedIndexChanged

    End Sub
End Class
