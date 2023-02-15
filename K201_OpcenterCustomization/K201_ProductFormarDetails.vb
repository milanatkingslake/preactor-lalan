Imports System.Data.SqlClient
Imports System.Reflection.Emit
Imports Preactor
Imports Preactor.Interop.PreactorObject


Public Class K201_ProductFormarDetails
    Public Property tblFormerDetailsMain As DataTable
    Public Property tblSize As DataTable
    Public Property tblOrder As DataTable
    Public Property connetionString As String
    Public Property lineType As String
    Public Property tbltblOrderRate_gl As DataTable

    Dim tblOrderRate As DataTable = New DataTable()


    Private Sub K201_ProductDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim orderCount As Integer = tblOrder.Rows.Count

        If lineType = "C" Then
            btnCalculatePlatoonRatio.Hide()
        End If


        DataGridViewCoProduct.DataSource = tblFormerDetailsMain
        DataGridViewCoProduct.Refresh()
        DataGridViewCoProduct.AutoResizeColumns()
        Me.DataGridViewCoProduct.Columns("OrderID").Visible = False
        Me.DataGridViewCoProduct.Rows(CInt(tblOrder.Rows.Count) + 6).ReadOnly = True

        Dim totalPlatoon As Decimal
        For Each size As DataRow In tblSize.Rows
            Dim size_s As String = size("Product").ToString()
            Dim allocatedPlatoons As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 7)(size_s).ToString)
            totalPlatoon = totalPlatoon + allocatedPlatoons
        Next

        DataGridViewCoProduct.Rows(orderCount + 7).Cells("Total").Value = totalPlatoon

        Try
            If lineType = "P" Then
                Me.DataGridViewCoProduct.Rows(CInt(tblOrder.Rows.Count) + 5).Visible = False
                Me.DataGridViewCoProduct.Rows(CInt(tblOrder.Rows.Count) + 9).Visible = False
            End If
        Catch ex As Exception
            MsgBox("Error in hiding the rows" & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub btnConfirmResourceRate_Click(sender As Object, e As EventArgs) Handles btnConfirmResourceRate.Click
        Dim confirmMsg As Integer

        Dim orderId As DataColumn = New DataColumn("OrderID", Type.[GetType]("System.Double"))
        Dim orderNum As DataColumn = New DataColumn("OrderNum", Type.[GetType]("System.String"))
        Dim rate As DataColumn = New DataColumn("Rate", Type.[GetType]("System.Double"))
        Dim formarRatioVal As DataColumn = New DataColumn("FormarRatio", Type.[GetType]("System.Double"))
        Dim totalFormers As DataColumn = New DataColumn("TotalFormers", Type.[GetType]("System.Double"))


        Dim orderCount As Integer = tblOrder.Rows.Count
        Dim total As String = "Total"
        Dim finalLineCapacity As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 4)(total).ToString)

        Dim totalQuantity As Decimal = 0

        For Each size As DataRow In tblSize.Rows
            Dim size_s As String = size("Product").ToString()
            Dim formarRatio As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 7)(size_s).ToString)
            totalQuantity = totalQuantity + formarRatio
        Next

        If totalQuantity <= finalLineCapacity And totalQuantity > 0 Then

            If tblOrderRate.Columns.Count > 0 Then
                MsgBox("Rate calculation already confirm",, "error")
            Else
                tblOrderRate.Columns.Add(orderId)
                tblOrderRate.Columns.Add(orderNum)
                tblOrderRate.Columns.Add(rate)
                tblOrderRate.Columns.Add(formarRatioVal)
                tblOrderRate.Columns.Add(totalFormers)


                confirmMsg = MsgBox("Do You Want To Confirm The Recalculation?", vbOKCancel, "Preactor Former Ratio")
                If confirmMsg = 1 Then
                    ''MsgBox("You have clicked the yes button")


                    For Each size As DataRow In tblSize.Rows
                        ''  MsgBox(size("Product").ToString())
                        Dim size_s As String = size("Product").ToString()
                        Dim orderResourceRate As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 9)(size_s).ToString)
                        Dim formarRatio As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 7)(size_s).ToString)
                        Dim totalFormer As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 6)("Total").ToString)

                        If orderResourceRate > 0 Then
                            Dim oc As Integer = orderCount
                            For Each order As DataRow In tblFormerDetailsMain.Rows
                                If oc > 0 Then
                                    If Not order(size_s).ToString() = "" Then
                                        Dim qt As Decimal = CDec(order(size_s).ToString())
                                        If qt > 0 Then
                                            Dim orderVal As String = order("#").ToString()
                                            Dim orderIdVal As Decimal = CDec(order("OrderId").ToString)

                                            Dim dt_sr As DataRow = tblOrderRate.NewRow()
                                            dt_sr("OrderID") = orderIdVal
                                            dt_sr("OrderNum") = orderVal
                                            dt_sr("Rate") = orderResourceRate
                                            dt_sr("FormarRatio") = formarRatio
                                            dt_sr("TotalFormers") = totalFormer

                                            tblOrderRate.Rows.Add(dt_sr)

                                        End If
                                    End If

                                End If
                                oc = oc - 1
                            Next
                        End If

                    Next

                    tbltblOrderRate_gl = tblOrderRate

                End If

            End If

        Else
            MsgBox("Total former quantity should not be more than Line Capacity",, "Please Enter the valid former quantity")
        End If

    End Sub

    Public Function K201_UpdateOrderResourceRate(ByRef connetionString As String, ByRef orerderNo As String, ByRef rate As Decimal) As Integer

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_UpdateOrdersResourceRate_Sp"
            Dim param As SqlParameter

            param = New SqlParameter("@OrderNum", orerderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@Rate", rate)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            Dim status As Decimal = 0
            param = New SqlParameter("@Status", status)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Boolean
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "1") Then
                Return 1
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Orders resource rate not define",, "error")
            ''MsgBox(ex.Message)
        Finally

        End Try

    End Function

    Private Sub K201_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        sender = tblOrderRate
    End Sub

    Private Sub btnCalculatePlatoonRatio_Click(sender As Object, e As EventArgs) Handles btnCalculatePlatoonRatio.Click

        Dim orderCount As Integer = tblOrder.Rows.Count
        Dim total As String = "Total"
        Dim finalLineCapacity As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 4)(total).ToString)
        Dim totalProducedGlovesInPair As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 11)(total).ToString)

        Dim totalQuantity As Decimal = 0
        Dim totalAllocatedPlatoons As Decimal
        Dim totalProducedGloveInPair As Decimal

        For Each size As DataRow In tblSize.Rows
            Dim size_s As String = size("Product").ToString()
            Dim totalPlatoon As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 3)(size_s).ToString)
            Dim allocatedPlatoons As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 7)(size_s).ToString)
            Dim formersPerPlatoon As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 10)(size_s).ToString)
            Dim producedGloveInPairs As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 11)(size_s).ToString)
            Dim totalRatio As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 1)(size_s).ToString)
            Dim producedSizeRatio As Decimal = CDec(tblFormerDetailsMain.Rows(orderCount + 12)(size_s).ToString)
            DataGridViewCoProduct.Rows(orderCount + 8).Cells(size_s).Value = totalPlatoon - allocatedPlatoons
            DataGridViewCoProduct.Rows(orderCount + 11).Cells(size_s).Value = (allocatedPlatoons * formersPerPlatoon) / 2
            DataGridViewCoProduct.Rows(orderCount + 12).Cells(size_s).Value = Math.Round((producedGloveInPairs / totalProducedGlovesInPair) * 100, 2)
            DataGridViewCoProduct.Rows(orderCount + 13).Cells(size_s).Value = totalRatio - producedSizeRatio
            totalAllocatedPlatoons = totalAllocatedPlatoons + CDec(tblFormerDetailsMain.Rows(orderCount + 7)(size_s).ToString)
            totalProducedGloveInPair = totalProducedGloveInPair + CDec(tblFormerDetailsMain.Rows(orderCount + 11)(size_s).ToString)
        Next
        DataGridViewCoProduct.Rows(orderCount + 7).Cells(total).Value = totalAllocatedPlatoons
        DataGridViewCoProduct.Rows(orderCount + 11).Cells(total).Value = totalProducedGloveInPair





    End Sub

    Private Sub DataGridViewCoProduct_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewCoProduct.CellContentClick

    End Sub
End Class