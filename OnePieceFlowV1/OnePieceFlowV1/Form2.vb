Imports System.Data.SqlClient
Public Class Form2
    Public Station

    Dim con As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Station " + Station
        Dim Labels As List(Of Label) = Form1.Controls.OfType(Of Label).ToList
        Label2.Text = Form1.TransferCase
        Dim i = 2
        ComboBox1.Items.Clear()

        ComboBox1.DisplayMember = "Text"
        ComboBox1.ValueMember = "Value"
        Dim tb As New DataTable
        tb.Columns.Add("Text", GetType(String))
        tb.Columns.Add("Value", GetType(Integer))

        If Station = 1 Then
            tb.Rows.Add("Remove Case", 0)
        Else
            tb.Rows.Add("Station 1", 0)
            Do While i < Station

                tb.Rows.Add("Station " + i.ToString, i)
                i = i + 1
            Loop
        End If


        ComboBox1.DataSource = tb

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Form2_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        Me.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (ComboBox1.Text <> "") Then
            If ComboBox1.Text <> Label1.Text Then
                If MessageBox.Show("Are you sure you want to Transfer this case from Station " + Station + " To " + ComboBox1.SelectedValue.ToString + "?", "Case Transfer", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                    If (ComboBox1.SelectedValue = 0) Then

                        con.Open()
                        Dim UpdateDetails As String = "DELETE From [SMProduction].[dbo].[ProductionDetails] Where ProductionDetailID IN (SELECT ProductionDetailID FROM [SMProduction].[dbo].[ProductionDetails] as PD LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID WHERE PH.SomtrackID = @Som)"
                        Dim UpdateDetailsQuery As SqlCommand = New SqlCommand(UpdateDetails, con)
                        UpdateDetailsQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateDetailsQuery.ExecuteNonQuery()
                        con.Close()

                        ''''UPDATE HEAD''''
                        con.Open()
                        Dim UpdateHead As String = "DELETE FROM ProductionHead WHERE SomtrackID = @Som"
                        Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con)
                        UpdateHeadQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateHeadQuery.ExecuteNonQuery()
                        con.Close()
                        Label4.Text = Label2.Text
                        Label2.Text = ""
                        MessageBox.Show("Case transfered to " + ComboBox1.SelectedValue.ToString, "Success")
                        Me.Close()
                    Else
                        ''''UPDATE ACTIVE DETAILS''''
                        con.Open()
                        Dim UpdateDetails As String = "update PD SET PD.Status = 4, PD.DateEnded = GETDATE() From [SMProduction].[dbo].[ProductionHead] as PH Left Join ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID WHERE PH.SomtrackID = @Som And PD.Status = 1"
                        Dim UpdateDetailsQuery As SqlCommand = New SqlCommand(UpdateDetails, con)
                        UpdateDetailsQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateDetailsQuery.ExecuteNonQuery()
                        con.Close()


                        ''''ZERO OUT INCLUDED DETAILS''''
                        con.Open()
                        Dim UpdateDetail As String = "UPDATE PD Set PD.Status = 4 FROM [SMProduction].[dbo].[Station] as S LEFT JOIN StationProcess as SP ON SP.StationID = S.StationID LEFT JOIN ProductionDetails as PD ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = Pd.ProductionHeadID WHERE (SP.StationID Between @S2 AND @S1) AND PH.SomtrackID = @Som"
                        Dim UpdateDetailQuery As SqlCommand = New SqlCommand(UpdateDetail, con)
                        UpdateDetailQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateDetailQuery.Parameters.AddWithValue("@S1", Station)
                        UpdateDetailQuery.Parameters.AddWithValue("@S2", ComboBox1.SelectedValue)
                        UpdateDetailQuery.ExecuteNonQuery()
                        con.Close()


                        ''''INSERT DETAILS''''
                        con.Open()
                        Dim InsertDetails As String = "INSERT INTO [ProductionDetails] ([ProductionHeadID], [BOMDID], [Status]) SELECT [ProductionHeadID], SP.BOMDID, 3 as Status FROM [ProductionHead] as PH LEFT JOIN Converter AS C ON C.ProductSubID = PH.ProductSubID LEFT JOIN BillOfMaterials as BM ON BM.BOMID = C.BOMID LEFT JOIN BillOfMaterialsDetails as BD ON BD.BOMID = BM.BOMID LEFT JOIN StationProcess as SP ON SP.BOMDID = BD.BOMDID WHERE BM.BOMStatus = 1 AND BD.BOMDStatus = 1 AND PH.SomtrackID = @Som AND SP.StationID BETWEEN @S2 AND @S1"
                        Dim InsertDetailsQuery As SqlCommand = New SqlCommand(InsertDetails, con)
                        InsertDetailsQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        InsertDetailsQuery.Parameters.AddWithValue("@S1", Station)
                        InsertDetailsQuery.Parameters.AddWithValue("@S2", ComboBox1.SelectedValue)
                        InsertDetailsQuery.ExecuteNonQuery()
                        con.Close()

                        ''''UPDATE HEAD''''
                        con.Open()
                        Dim UpdateHead As String = "UPDATE ProductionHead SET StationID = @SID WHERE SomtrackID = @Som"
                        Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con)
                        UpdateHeadQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateHeadQuery.Parameters.AddWithValue("@SID", ComboBox1.SelectedValue)
                        UpdateHeadQuery.ExecuteNonQuery()
                        con.Close()


                        ''''UPDATE NEXT DETAILS''''
                        con.Open()
                        Dim UpdateNextDetails As String = "Update PD SET PD.Status = 2 FROM [SMProduction].[dbo].[ProductionHead] as PH LEFT JOIN StationProcess as SP ON SP.StationID = PH.StationID LEFT JOIN ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID And PD.BOMDID = SP.BOMDID WHERE PH.SomtrackID = @Som AND PD.Status = 3"
                        Dim UpdateNextDetailsQuery As SqlCommand = New SqlCommand(UpdateNextDetails, con)
                        UpdateNextDetailsQuery.Parameters.AddWithValue("@Som", Label2.Text)
                        UpdateNextDetailsQuery.ExecuteNonQuery()
                        con.Close()

                        Label4.Text = Label2.Text
                        Label2.Text = ""
                        MessageBox.Show("Case transfered to " + ComboBox1.SelectedValue.ToString, "Success")
                        Me.Close()
                    End If
                End If
                Else
                MessageBox.Show("Cannot transfer to same station", "Error")
            End If
        Else
            MessageBox.Show("Please select receiving station", "Error")
        End If
    End Sub
End Class
