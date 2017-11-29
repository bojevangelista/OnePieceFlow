Imports System.Data.SqlClient
Public Class Form1
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\Programmer\Documents\OPF-TableNo.txt")
    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            Dim con As SqlConnection
            Label15.Text = TextBox1.Text
            TextBox1.Text = ""
            'LOCALHOST'
            'con = New SqlConnection("Data Source=localhost;Initial Catalog=SomnoMed;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
            'LOCALHOST'

            'SERVER'
            con = New SqlConnection("Data Source=10.130.15.40;Initial Catalog=somtrackdbprod;User ID=somtrack2;Password=sompass12345")


            '''''''QUERY FOR SELECTING PROCESS OF THE PRODUCT'''''''''''
            Dim prodquery As String = "Select LDT.DeviceTypeName as DeviceName From LstDevice as LD Left Join LstDeviceType as LDT ON LD.ProductTypeID = LDT.DeviceTypeId WHERE LD.DeviceID = @prod"
            Dim prodcmd As SqlCommand = New SqlCommand(prodquery, con)
            prodcmd.Parameters.AddWithValue("@prod", Label15.Text)
            con.Open()
            Using reader As SqlDataReader = prodcmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read()
                        Label3.Text = reader.Item("DeviceName")

                    End While

                Else

                    Label3.Text = "Invalid Somtrack"

                End If
            End Using


            con.Close()

        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If (CheckBox1.Checked = True) Then
            Panel1.Enabled = True
            Label7.Enabled = True
        Else
            Panel1.Enabled = False
            Label7.Enabled = False
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim familyQuery As String = "SELECT ProductFamilyID, ProductFamilyName FROM ProductFamily"
        Dim familyCmd As SqlCommand = New SqlCommand(familyQuery, con)
        con.Open()
        Using reader As SqlDataReader = familyCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox1.DataSource = dt
                ComboBox1.ValueMember = "ProductFamilyID"
                ComboBox1.DisplayMember = "ProductFamilyName"

            End If
        End Using
        con.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim CB1 As String = ComboBox1.SelectedValue.ToString
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim typeQuery As String = "SELECT [ProductTypeID], [ProductTypeName] FROM [ProductType] WHERE [ProductFamilyID] = @PF"
        Dim typeCmd As SqlCommand = New SqlCommand(typeQuery, con)
        typeCmd.Parameters.AddWithValue("@PF", CB1)
        con.Open()
        Using reader As SqlDataReader = typeCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox2.DataSource = dt
                ComboBox2.ValueMember = "ProductTypeID"
                ComboBox2.DisplayMember = "ProductTypeName"

            End If
        End Using
        con.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        '''''''VARIABLES'''''''
        Dim con As SqlConnection
        Dim CB2 As String = ComboBox2.SelectedValue.ToString
        'SERVER'
        con = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
        '''''''QUERY FOR SELECTING LEADMAN'''''''''''
        Dim subQuery As String = "SELECT [ProductSubID], CONCAT([ProductSubTypeName], ' ',[ProductSubClass]) as Subname  FROM [ProductSubType] WHERE [ProductTypeID] = @PT"
        Dim subCmd As SqlCommand = New SqlCommand(subQuery, con)
        subCmd.Parameters.AddWithValue("@PT", CB2)
        con.Open()
        Using reader As SqlDataReader = subCmd.ExecuteReader()
            If reader.HasRows Then
                Dim dt As DataTable = New DataTable
                dt.Load(reader)

                ComboBox3.DataSource = dt
                ComboBox3.ValueMember = "ProductSubID"
                ComboBox3.DisplayMember = "Subname"

            End If
        End Using
        con.Close()
    End Sub
End Class
