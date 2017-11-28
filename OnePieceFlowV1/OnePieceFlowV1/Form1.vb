Imports System.Data.SqlClient


Public Class Form1
    Dim ActiveStation = 0
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\Programmer\Documents\OPF-TableNo.txt")

    Private Sub ChangeEmployee(e, f, g)
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim cEmpquery As String = "SELECT [Username] FROM [Users] WHERE Username = @emp"
        Dim cEmpcmd As SqlCommand = New SqlCommand(cEmpquery, con)
        cEmpcmd.Parameters.AddWithValue("@emp", e.SelectedValue.ToString)
        con.Open()
        Using reader As SqlDataReader = cEmpcmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    Dim Hours, Minutes, Seconds As Integer
                    Seconds = Integer.Parse(CInt(Math.Ceiling(Rnd() * 500)) + 1)
                    Hours = Seconds / 3600
                    Seconds = Seconds Mod 3600
                    Minutes = Seconds / 60
                    Seconds = Seconds Mod 60

                    f.Text = Hours.ToString.PadLeft(2, "0"c) & ":" & Minutes.ToString.PadLeft(2, "0"c) & ":" & Seconds.ToString.PadLeft(2, "0"c)
                    g.Text = "Timed In"
                End While

            End If
        End Using
        con.Close()
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            If (TextBox1.Text <> "") Then
                TextBox1.Visible = False
                Label21.Visible = True
                Label21.Text = TextBox1.Text
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetLeadman()
        GetTableDetails()

    End Sub

    Private Sub Clock_Tick(sender As Object, e As EventArgs) Handles Clock.Tick
        Label1.Text = Format(Now, "hh:mm:ss")
        Label2.Text = Format(Now, "MMMM dd, yyyy")
    End Sub
    Private Sub GetTableDetails()
        '''''''VARIABLES'''''''
        Dim con As SqlConnection

        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING ROSTER'''''''''''
        Dim empquery As String = "SELECT [EmployeeID] ,CONCAT([LastName],', ',[FirstName], ' - ', [Username]) as FullName FROM [SMProduction].[dbo].[TableMembers] as TM LEFT JOIN [SMProduction].[dbo].[TableSet] as TS ON TS.TableSetID = TM.TableSetID LEFT JOIN [SMBuildLog].[dbo].[Users] as U ON TM.EmployeeID = U.Username WHERE TS.TableID = @TN AND TableMemberStatus = 1 AND StationID = @SN AND TS.TableSetStatus = 1"
        Dim empcmd As SqlCommand = New SqlCommand(empquery, con)


        Dim empquery2 As String = "SELECT '0' as Username, '' as Fullname UNION SELECT [Username] ,CONCAT([LastName],', ',[FirstName], ' - ',[Username]) as Fullname FROM [Users] WHERE DepID NOT IN (9,8,13,14,11,10,15,16,17,18) AND Status = 'Active' ORDER BY FullName"
        Dim empcmd2 As SqlCommand = New SqlCommand(empquery2, con)


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
                    con.Open()

                    Using reader2 As SqlDataReader = empcmd2.ExecuteReader()
                        Dim dt2 As DataTable = New DataTable
                        dt2.Load(reader2)

                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DataSource = dt2
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).ValueMember = "Username"
                        CType(CType(Me.Controls("Panel" + (i - 1).ToString()), Panel).Controls("ComboBox" + (i).ToString()), ComboBox).DisplayMember = "FullName"

                    End Using
                    con.Close()
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
        leadcurrcmd.Parameters.AddWithValue("@TN", 1)

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
            con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

            '''''''QUERY FOR SELECTING LEADMAN'''''''''''
            Dim empquery As String = "SELECT '0' as Username, '' as Fullname UNION SELECT [Username] ,CONCAT([LastName],', ',[FirstName], ' - ',[Username]) as Fullname FROM [Users] WHERE DepID NOT IN (9,8,13,14,11,10,15,16,17,18) AND Status = 'Active' ORDER BY FullName"
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
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim empquery As String = "SELECT '0' as Username, '' as Fullname UNION SELECT [Username] ,CONCAT([LastName],', ',[FirstName], ' - ',[Username]) as Fullname FROM [Users] WHERE DepID NOT IN (9,8,13,14,11,10,15,16,17,18) AND Status = 'Active' ORDER BY FullName"
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


                If MessageBox.Show("Are you sure you want to lock in this roster?", "Title", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                    If (ActiveStation = 12) Then


                        '''''''VARIABLES'''''''
                        Dim con As SqlConnection
                        'SERVER'
                        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
                        con.Open()

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
                        MessageBox.Show("Please fill all the station")
                    End If

                End If
            End If
        Else
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
        TextBox1.Visible = True
        Label21.Visible = False
        TextBox1.Text = Label21.Text
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label79_Click(sender As Object, e As EventArgs) Handles Label79.Click

    End Sub

    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click

    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
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

            ComboBox1.Enabled = True

        End If

    End Sub
End Class
