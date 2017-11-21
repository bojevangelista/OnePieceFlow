Imports System.Data.SqlClient
Public Class Form1
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

                    g.Text = Hours.ToString.PadLeft(2, "0"c) & ":" & Minutes.ToString.PadLeft(2, "0"c) & ":" & Seconds.ToString.PadLeft(2, "0"c)
                    f.Text = "Timed In"
                End While

            End If
        End Using
        con.Close()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetLeadman()
        GetEmployee()

    End Sub

    Private Sub Clock_Tick(sender As Object, e As EventArgs) Handles Clock.Tick
        Label1.Text = Format(Now, "hh:mm:ss")
        Label2.Text = Format(Now, "MMMM dd, yyyy")
    End Sub

    Private Sub GetLeadman()
        '''''''VARIABLES'''''''
        Dim con As SqlConnection

        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMBuildLog;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim leadquery As String = "SELECT '0' as Username, '' as FullName UNION SELECT [Username] ,CONCAT([LastName],', ',[FirstName], ' - ', [Username]) as FullName FROM [Users] WHERE DepID = '9' AND Status = 'Active' ORDER BY FullName"
        Dim leadcmd As SqlCommand = New SqlCommand(leadquery, con)
        ComboBox1.Items.Insert(0, String.Empty)
        con.Open()
        Using reader As SqlDataReader = leadcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)
            ComboBox1.DataSource = dt
            ComboBox1.ValueMember = "Username"
            ComboBox1.DisplayMember = "FullName"

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
            End If
        Else
            x.Text = "Sign In"
            x.BackColor = Color.LightGray
            e.Enabled = True
            f.BackColor = Color.DarkGray
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
        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox2.DataSource = dt
            ComboBox2.ValueMember = "Username"
            ComboBox2.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox3.DataSource = dt
            ComboBox3.ValueMember = "Username"
            ComboBox3.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox4.DataSource = dt
            ComboBox4.ValueMember = "Username"
            ComboBox4.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox5.DataSource = dt
            ComboBox5.ValueMember = "Username"
            ComboBox5.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox6.DataSource = dt
            ComboBox6.ValueMember = "Username"
            ComboBox6.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox7.DataSource = dt
            ComboBox7.ValueMember = "Username"
            ComboBox7.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox8.DataSource = dt
            ComboBox8.ValueMember = "Username"
            ComboBox8.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox9.DataSource = dt
            ComboBox9.ValueMember = "Username"
            ComboBox9.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox10.DataSource = dt
            ComboBox10.ValueMember = "Username"
            ComboBox10.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox11.DataSource = dt
            ComboBox11.ValueMember = "Username"
            ComboBox11.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox12.DataSource = dt
            ComboBox12.ValueMember = "Username"
            ComboBox12.DisplayMember = "FullName"
        End Using

        Using reader As SqlDataReader = empcmd.ExecuteReader()

            Dim dt As DataTable = New DataTable
            dt.Load(reader)

            ComboBox13.DataSource = dt
            ComboBox13.ValueMember = "Username"
            ComboBox13.DisplayMember = "FullName"
        End Using

        con.Close()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Button1.Text = "On Break"
        Button1.BackColor = Color.Orange

        Button2.Text = "On Break"
        Button2.BackColor = Color.Orange

        Button3.Text = "On Break"
        Button3.BackColor = Color.Orange

        Button4.Text = "On Break"
        Button4.BackColor = Color.Orange

        Button5.Text = "On Break"
        Button5.BackColor = Color.Orange

        Button6.Text = "On Break"
        Button6.BackColor = Color.Orange

        Button7.Text = "On Break"
        Button7.BackColor = Color.Orange

        Button8.Text = "On Break"
        Button8.BackColor = Color.Orange

        Button9.Text = "On Break"
        Button9.BackColor = Color.Orange

        Button10.Text = "On Break"
        Button10.BackColor = Color.Orange

        Button11.Text = "On Break"
        Button11.BackColor = Color.Orange

        Button12.Text = "On Break"
        Button12.BackColor = Color.Orange




    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click


        Button1.Text = "Break Time"
        Button1.BackColor = Color.LawnGreen

        Button2.Text = "Break Time"
        Button2.BackColor = Color.LawnGreen

        Button3.Text = "Break Time"
        Button3.BackColor = Color.LawnGreen

        Button4.Text = "Break Time"
        Button4.BackColor = Color.LawnGreen

        Button5.Text = "Break Time"
        Button5.BackColor = Color.LawnGreen

        Button6.Text = "Break Time"
        Button6.BackColor = Color.LawnGreen

        Button7.Text = "Break Time"
        Button7.BackColor = Color.LawnGreen

        Button8.Text = "Break Time"
        Button8.BackColor = Color.LawnGreen

        Button9.Text = "Break Time"
        Button9.BackColor = Color.LawnGreen

        Button10.Text = "Break Time"
        Button10.BackColor = Color.LawnGreen

        Button11.Text = "Break Time"
        Button11.BackColor = Color.LawnGreen

        Button12.Text = "Break Time"
        Button12.BackColor = Color.LawnGreen
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
        SignStatus(Button17, ComboBox5, Panel2)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        SignStatus(Button15, ComboBox3, Panel4)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        SignStatus(Button18, ComboBox6, Panel3)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        SignStatus(Button19, ComboBox7, Panel5)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        SignStatus(Button16, ComboBox4, Panel6)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        SignStatus(Button20, ComboBox8, Panel12)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        SignStatus(Button21, ComboBox9, Panel10)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        SignStatus(Button22, ComboBox10, Panel11)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        SignStatus(Button23, ComboBox11, Panel8)
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        SignStatus(Button25, ComboBox13, Panel7)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        SignStatus(Button24, ComboBox12, Panel9)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If (Button13.Text = "Lock In") Then
            If (ComboBox1.Text = "") Then
                MessageBox.Show("Leadman cannot be empty", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
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
            End If
        Else
            Button13.BackColor = Color.LightGray
            Button13.Text = "Lock In"
            ComboBox1.Enabled = True
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
        End If

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ChangeEmployee(ComboBox2, Label20, Label18)
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ChangeEmployee(ComboBox3, Label22, Label24)
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ChangeEmployee(ComboBox4, Label27, Label29)
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        ChangeEmployee(ComboBox8, Label47, Label49)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ChangeEmployee(ComboBox5, Label32, Label34)
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ChangeEmployee(ComboBox6, Label37, Label39)
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ChangeEmployee(ComboBox7, Label42, Label44)
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        ChangeEmployee(ComboBox9, Label52, Label54)
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        ChangeEmployee(ComboBox10, Label57, Label59)
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        ChangeEmployee(ComboBox11, Label62, Label64)
    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        ChangeEmployee(ComboBox13, Label72, Label74)
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        ChangeEmployee(ComboBox12, Label67, Label69)
    End Sub
End Class
