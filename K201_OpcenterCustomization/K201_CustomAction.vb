Option Strict On
Option Explicit On
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)>
<Microsoft.VisualBasic.ComClass("611195b3-9f96-43ed-afeb-e97b3a9ce91e", "10aa0b6a-97d9-486e-a5bf-9c563953a9c5")>
Public Class K201_CustomAction
#Region "Co_Product_Form_Save"
    ''Co-Product window saving
    Public Function Co_Product_Form_Save(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef i As Integer) As Integer

        Try


            Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
            preactor.ReadFieldString("Co-products", "Co-product", i)
            Dim strOrderNo As String = preactor.ReadFieldString("Co-products", "Order No.", i)
            Dim strCoProductName As String = preactor.ReadFieldString("Co-products", "Co-product", i)
            Dim strcoProductQty As String = preactor.ReadFieldString("Co-products", "Quantity", i)
            Dim OpNum As Integer = CInt(preactor.ReadFieldString("Co-products", "Op. No.", i))
            ''Order Record Number
            Dim orderRecordNum As Integer = preactor.FindMatchingRecord("Orders", "Order No.", orderRecordNum, strOrderNo)

            Dim coProductRecordNum As Integer = 0
            Dim intMaxSeq As Decimal = New Decimal()
            Dim intCurrentSeq As Decimal = New Decimal()
            ''get co-product record Number
            coProductRecordNum = preactor.FindMatchingRecord("Co-products", "Order No.", coProductRecordNum, strOrderNo)
            ''check "K201_Order_Display_Sequence" feild is empty
            If (Not Information.IsNothing(preactor.ReadFieldString("Co-products", "K201_Order_Display_Sequence", coProductRecordNum))) Then
                intCurrentSeq = CDec(preactor.ReadFieldString("Co-products", "K201_Order_Display_Sequence", coProductRecordNum))
                ''Assign co product sequnce and quantities
                While orderRecordNum > 0
                    If OpNum = CInt(preactor.ReadFieldString("Orders", "Op. No.", orderRecordNum)) Then
                        If intCurrentSeq = 1 Then
                            preactor.WriteField("Orders", "K201 Co Product 1", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 1", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 1", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 2 Then
                            preactor.WriteField("Orders", "K201 Co Product 2", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 2", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 2", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 3 Then
                            preactor.WriteField("Orders", "K201 Co Product 3", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 3", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 3", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 4 Then
                            preactor.WriteField("Orders", "K201 Co Product 4", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 4", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 4", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 5 Then
                            preactor.WriteField("Orders", "K201 Co Product 5", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 5", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 5", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 6 Then
                            preactor.WriteField("Orders", "K201 Co Product 6", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 6", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 6", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 7 Then
                            preactor.WriteField("Orders", "K201 Co Product 7", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 7", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 7", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 8 Then
                            preactor.WriteField("Orders", "K201 Co Product 8", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 8", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 8", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 9 Then
                            preactor.WriteField("Orders", "K201 Co Product 9", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 9", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 9", orderRecordNum, strcoProductQty)
                        ElseIf intCurrentSeq = 10 Then
                            preactor.WriteField("Orders", "K201 Co Product 10", orderRecordNum, strCoProductName)
                            preactor.WriteField("Orders", "K201 Co Product Sequance 10", orderRecordNum, intCurrentSeq)
                            preactor.WriteField("Orders", "K201 Co Product Quantity 10", orderRecordNum, strcoProductQty)
                        End If

                    End If

                    orderRecordNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRecordNum, strOrderNo)
                End While

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        '' preactor.WriteField("Co-products", "CoProduct Sequence Number", i, Convert.ToDouble(Decimal.Add(intMaxSeq, Decimal.One)))
        Return 0
    End Function
#End Region
    ''In Gant chart if user selected block drag this event will execute 
#Region "Gant Chart Drag and Drop"
    Public Function K201_GatChartBlockDrag(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef i As Integer) As Integer
        '' define variable and assign
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim setup_Start As Date = preactor.ReadFieldDateTime("Orders", "Setup Start", i)
        Dim resource As String = preactor.ReadFieldString("Orders", "Resource", i)
        Dim intrecnum As Integer = preactor.CreateRecord("dragtemp")
        ''Asign "dragtemp" table to resource and setup start time
        preactor.WriteField("dragtemp", "Resource", intrecnum, resource)
        preactor.WriteField("dragtemp", "Setupstart", intrecnum, setup_Start)
        preactor.Commit("dragtemp")

        Return 0
    End Function
    ''In Gant chart if user selected block drop this event will execute
    Public Function K201_GatChartBlockDrop(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef i As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        ''get auto split varible value
        Dim isAutoSplit As Boolean = preactor.ReadFieldBool("Auto Split Switch", "Toggle", 1)
        ''check auto split varible value
        If isAutoSplit Then
            '' define variable and assign
            Dim setup_Start As Date = preactor.ReadFieldDateTime("Orders", "Setup Start", i)
            Dim end_Time As Date = preactor.ReadFieldDateTime("Orders", "End Time", i)
            Dim drag_resource As String = preactor.ReadFieldString("Orders", "Resource", i)
            Dim blockQty As Double = preactor.ReadFieldDouble("Orders", "Quantity", i)
            Dim drop_resource As String = preactor.ReadFieldString("dragtemp", "Resource", 1)
            Dim pb As IPlanningBoard = preactor.PlanningBoard
            ''Get previos operation id
            Dim previousOperation As Integer = pb.GetResourcesPreviousOperation(pb.GetResourceNumber(drag_resource), setup_Start)
            Dim previousOperationEndTime As Date
            Dim NextOperationSetupStartTime As Date

            If previousOperation > 0 Then
                previousOperationEndTime = preactor.ReadFieldDateTime("Orders", "End Time", previousOperation)

                Dim NextOperation As Integer = pb.GetResourcesNextOperation(pb.GetResourceNumber(drag_resource), previousOperationEndTime)

                If (NextOperation <> i) And (NextOperation > 0) Then
                    NextOperationSetupStartTime = preactor.ReadFieldDateTime("Orders", "Setup Start", NextOperation)

                    MsgBox(previousOperationEndTime.ToString + "===" + NextOperationSetupStartTime.ToString)

                    pb.PutOperationOnResource(i, pb.GetResourceNumber(drag_resource), previousOperationEndTime)
                    Dim splitqty As Double = pb.GetProcessedQuantity(pb.GetResourceNumber(drag_resource), i, previousOperationEndTime, NextOperationSetupStartTime)
                    ''worte split quantity to Order table
                    preactor.WriteField("Orders", "Quantity", i, splitqty)
                    '' Recalculate Operation Time
                    pb.RecalculateOperationTimes(pb.GetResourceNumber(drag_resource))
                    ''assign resource to operation
                    pb.PutOperationOnResource(NextOperation, pb.GetResourceNumber(drag_resource), NextOperationSetupStartTime)
                    ''get new block quantity
                    Dim newBlockQty As Double = (blockQty - splitqty)
                    If newBlockQty > 0 Then

                        Dim newBlock As Integer = preactor.CreateRecord("Orders")
                        Dim newRecordNum As Integer = preactor.ReadFieldInt("Orders", "Number", newBlock)
                        preactor.CopyRecord("Orders", i, newBlock)
                        preactor.WriteField("Orders", "Number", newBlock, newRecordNum)
                        preactor.WriteField("Orders", "Quantity", newBlock, newBlockQty)

                        ''Check Resource and Set hold the order
                        Dim resourceName As String = preactor.ReadFieldString("Orders", "Resource", newBlock)
                        Dim op_no As Integer = preactor.ReadFieldInt("Orders", "Op. No.", newBlock)

                        If Not ((resourceName = "Nothing") Or (resourceName = "Unspecified")) Then
                            Dim K201_ProductionLineOperationRecordId As Integer = preactor.FindMatchingRecord("K201_ProductionLineOperation", "Line", K201_ProductionLineOperationRecordId, resourceName)

                            If K201_ProductionLineOperationRecordId > 0 Then

                                If op_no = 10 Then
                                    Dim Dipping As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Dipping", K201_ProductionLineOperationRecordId)
                                    If Not Dipping = 1 Then
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 1)
                                    Else
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 0)
                                    End If
                                End If
                                If op_no = 20 Then
                                    Dim Chlorination As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Chlorination", K201_ProductionLineOperationRecordId)
                                    If Not Chlorination = 1 Then
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 1)
                                    Else
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 0)
                                    End If
                                End If
                                If op_no = 30 Then
                                    Dim Printing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Printing", K201_ProductionLineOperationRecordId)
                                    If Not Printing = 1 Then
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 1)
                                    Else
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 0)
                                    End If
                                End If
                                If op_no = 40 Then
                                    Dim Packing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Packing", K201_ProductionLineOperationRecordId)
                                    If Not Packing = 1 Then
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 1)
                                    Else
                                        preactor.WriteField("Orders", "Disable Operation", newBlock, 0)
                                    End If
                                End If
                            End If
                        End If
                        preactor.Commit("Orders")
                        '''Check Resource and Set hold the order
                    End If
                    preactor.Redraw() ''referesh the gant chart
                End If
            End If
        End If
        Return 0
    End Function
#End Region

#Region "Export Order"
    '' export order strode procedure call
    Public Function K201_ExportOrders(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim orderCount As Integer = preactor.RecordCount("Orders")
        ''Get connection string 
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim connetionString_mod As String = connetionString
        connetionString_mod = connetionString_mod.Substring(connetionString_mod.IndexOf("="c) + 1)
        connetionString_mod = connetionString_mod.Substring(connetionString_mod.IndexOf("="c) + 1)
        Dim dbStrStart As String = connetionString_mod
        Dim dbEndIndex As Integer = connetionString_mod.IndexOf(";"c)
        Dim strOpDbName As String = dbStrStart.Substring(0, dbEndIndex)

        connetionString = connetionString.Replace(strOpDbName, "LALAN_intermediate_DB")

        Dim rowCount As Integer = 1
        Try
            Do
                ''define variable and assign
                Dim order_No As String = preactor.ReadFieldString("Orders", "Order No.", orderCount)
                Dim op_No As Integer = preactor.ReadFieldInt("Orders", "Op. No.", orderCount)
                Dim order_Start As Date = preactor.ReadFieldDateTime("Orders", "Order Start", orderCount)
                Dim order_End As Date = preactor.ReadFieldDateTime("Orders", "Order End", orderCount)
                Dim setup_Start As Date = preactor.ReadFieldDateTime("Orders", "Setup Start", orderCount)
                Dim start_Time As Date = preactor.ReadFieldDateTime("Orders", "Start Time", orderCount)
                Dim End_Time As Date = preactor.ReadFieldDateTime("Orders", "End Time", orderCount)

                Dim connection As SqlConnection
                Dim adapter As SqlDataAdapter
                Dim command As New SqlCommand

                connection = New SqlConnection(connetionString)
                ''execute K201_ExportOrders_Sp 
                connection.Open()
                command.Connection = connection
                command.CommandType = CommandType.StoredProcedure
                command.CommandText = "K201_ExportOrders_Sp"
                command.CommandTimeout = 600

                Dim param As SqlParameter

                param = New SqlParameter("@orderNo", order_No)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@oPNo", op_No)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@orderStart", order_Start)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@orderEnd", order_End)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@startTime", start_Time)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@EndTime", End_Time)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                param = New SqlParameter("@SetupStart", setup_Start)
                param.Direction = ParameterDirection.Input
                param.DbType = DbType.String
                command.Parameters.Add(param)

                adapter = New SqlDataAdapter(command)
                command.ExecuteNonQuery()

                connection.Close()
                rowCount = rowCount + 1

            Loop While rowCount <= orderCount

            MsgBox("Export Completed")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

        End Try

    End Function
#End Region

#Region "Former Ratio Calculation"
#Region "Former Ratio Calculation Co-Product"
    ''Get K201_CoProductFormerRatioCalculation 
    Public Function K201_CoProductFormerRatioCalculation(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim strFirstPartNo As String = ""
        Dim intSumQuanity As Decimal = New Decimal()
        Dim corProducts(preactor.RecordCount("Orders") + 1 - 1, 2) As String
        Dim decTotalOrderQuantity As Decimal = New Decimal()
        Dim dt As DataTable = New DataTable()
        ''define variable and assign
        Dim c_ProductNo As DataColumn = New DataColumn("ProductNo", Type.[GetType]("System.String"))
        Dim c_ProductQuantity As DataColumn = New DataColumn("ProductQuantity", Type.[GetType]("System.Double"))
        Dim c_OrderNo As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_OrderQuantity As DataColumn = New DataColumn("OrderQuantity", Type.[GetType]("System.String"))
        Dim c_ItemType As DataColumn = New DataColumn("ItemType", Type.[GetType]("System.String"))

        dt.Columns.Add(c_ProductNo)
        dt.Columns.Add(c_ProductQuantity)
        dt.Columns.Add(c_OrderNo)
        dt.Columns.Add(c_OrderQuantity)
        dt.Columns.Add(c_ItemType)

        Dim num As Integer = preactor.RecordCount("Orders")
        Dim i As Integer = 1
        Do
            If (planningboard.GetOperationLocateState(i)) Then
                Dim strPartNo As String = preactor.ReadFieldString("Orders", "Part No.", i)
                ''Get total quantity
                decTotalOrderQuantity = Decimal.Add(decTotalOrderQuantity, CDec(preactor.ReadFieldString("Orders", "Quantity", i)))
                If (strFirstPartNo = "") Then
                    strFirstPartNo = strPartNo
                End If
                If (Not Information.IsNothing(strPartNo)) Then
                    Dim dr_int As DataRow = dt.NewRow()
                    dr_int("ProductNo") = strPartNo
                    dr_int("ProductQuantity") = CDbl(preactor.ReadFieldString("Orders", "Quantity", i))
                    dr_int("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", i)
                    dr_int("OrderQuantity") = preactor.ReadFieldString("Orders", "Quantity", i)
                    dr_int("ItemType") = "M"
                    dt.Rows.Add(dr_int)
                End If
            End If
            i = i + 1
        Loop While i <= num

        dt.DefaultView.Sort = "ProductNo ASC, OrderNo ASC"

        dt = dt.DefaultView.ToTable()

        Dim dtCal As DataTable = New DataTable()
        Dim dt_s As DataTable = New DataTable()
        Dim c_CoProductNo_s As DataColumn = New DataColumn("CoProduct", Type.[GetType]("System.String"))
        Dim c_CoProductQuantity_s As DataColumn = New DataColumn("CoProductQuantity", Type.[GetType]("System.String"))
        Dim c_OrderNo_s As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_OrderQuantity_s As DataColumn = New DataColumn("OrderQuantity", Type.[GetType]("System.Double"))
        Dim c_ItemType_s As DataColumn = New DataColumn("ItemType", Type.[GetType]("System.String"))

        dt_s.Columns.Add(c_CoProductNo_s)
        dt_s.Columns.Add(c_CoProductQuantity_s)
        dt_s.Columns.Add(c_OrderNo_s)
        dt_s.Columns.Add(c_OrderQuantity_s)
        dt_s.Columns.Add(c_ItemType_s)

        Dim totalOrderQty As Decimal = 0
        For Each product_row As DataRow In dt.Rows
            Dim productNo As String = product_row("ProductNo").ToString()
            Dim OrderNo As String = product_row("OrderNo").ToString()
            Dim ProductQuantity As String = product_row("ProductQuantity").ToString()

            Dim coProductRecordNum As Integer = 0
            Dim intQty As Decimal = New Decimal()

            coProductRecordNum = preactor.FindMatchingRecord("Co-products", "Order No.", coProductRecordNum, OrderNo)
            Dim count As Integer = 1
            While coProductRecordNum > 0
                If (Not Information.IsNothing(preactor.ReadFieldString("Co-products", "Co-product", coProductRecordNum))) Then
                    intQty = Decimal.Add(intQty, CDec(preactor.ReadFieldString("Co-products", "Quantity", coProductRecordNum)))
                    Dim dt_sr As DataRow = dt_s.NewRow()

                    dt_sr("CoProduct") = preactor.ReadFieldString("Co-products", "Co-product", coProductRecordNum)
                    dt_sr("CoProductQuantity") = preactor.ReadFieldString("Co-products", "Quantity", coProductRecordNum)
                    dt_sr("OrderNo") = OrderNo
                    dt_sr("OrderQuantity") = ProductQuantity
                    dt_sr("ItemType") = "C"
                    dt_s.Rows.Add(dt_sr)
                    coProductRecordNum = preactor.FindMatchingRecord("Co-products", "Order No.", coProductRecordNum, OrderNo)
                    count = count + 1
                End If
            End While

        Next product_row
        ''data table shorting
        dt_s.DefaultView.Sort = "CoProduct ASC, OrderNo ASC"
        dt_s = dt_s.DefaultView.ToTable()

        Dim distinct_dt As DataTable = dt_s.DefaultView.ToTable(True, "CoProduct")

        Dim dtSumCal As DataTable = New DataTable()
        Dim c_SumCoProductNo_s As DataColumn = New DataColumn("CoProduct", Type.[GetType]("System.String"))
        Dim c_SumCoProductQuantity_s As DataColumn = New DataColumn("CoProductQuantity", Type.[GetType]("System.String"))

        dtSumCal.Columns.Add(c_SumCoProductNo_s)
        dtSumCal.Columns.Add(c_SumCoProductQuantity_s)

        For Each product_row As DataRow In distinct_dt.Rows
            Dim coProduct As String = product_row("CoProduct").ToString()
            Dim itemQty As Decimal = 0
            For Each inner_row As DataRow In dt_s.Rows
                If inner_row("CoProduct").ToString() = coProduct Then
                    itemQty = itemQty + CDec(inner_row("CoProductQuantity"))
                    totalOrderQty = totalOrderQty + CDec(inner_row("OrderQuantity"))
                End If
            Next inner_row

            Dim drSumCal_int As DataRow = dtSumCal.NewRow()
            drSumCal_int("CoProduct") = coProduct
            drSumCal_int("CoProductQuantity") = CDbl(itemQty)
            dtSumCal.Rows.Add(drSumCal_int)
        Next product_row

        Dim strOrderDetails As String = ""

        For Each product_row As DataRow In dtSumCal.Rows
            strOrderDetails = String.Concat(New String() {strOrderDetails, "" & vbCrLf & "Product =", product_row("CoProduct").ToString(), " Order Quantity =", product_row("CoProductQuantity").ToString(), " Former Ratio = ", Strings.Format(Decimal.Multiply(Decimal.Divide(CDec(product_row("CoProductQuantity").ToString()), totalOrderQty), New Decimal(CLng(100))), "0.00").ToString()})
        Next
        MsgBox(strOrderDetails, MsgBoxStyle.OkOnly, Nothing)
    End Function
    ''K201_FormerRatioCalculationCoProduct
    Public Function K201_FormerRatioCalculationCoProduct(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        ''Get connection string 
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")

        Dim strFirstPartNo As String = ""
        Dim intSumQuanity As Decimal = New Decimal()
        Dim corProducts(preactor.RecordCount("Orders") + 1 - 1, 2) As String
        Dim dt As DataTable = New DataTable()
        ''define variable and assign

        Dim c_OrderNo As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_OrderQuantity As DataColumn = New DataColumn("OrderQuantity", Type.[GetType]("System.String"))
        Dim c_OrderStartDate As DataColumn = New DataColumn("OrderStartDate", Type.[GetType]("System.String"))
        dt.Columns.Add(c_OrderNo)
        dt.Columns.Add(c_OrderQuantity)
        dt.Columns.Add(c_OrderStartDate)

        Dim line_capacity As Integer = 0
        Dim line_spreedPreDay As Integer = 1
        Dim pitch_Lenght As Decimal = 1
        Dim line_Speed As Decimal = 1
        Dim orderResourceRate As Decimal = 1

        Dim product As String = ""
        Dim damage_Percentage As Decimal = 1

        Dim wastage_Precentage As Double = 0.015
        Dim resource As String
        Dim jobStartDate As DateTime
        Dim numOfOrders As Integer = 0

        Dim num As Integer = preactor.RecordCount("Orders")
        Dim i As Integer = 1
        Do
            If (planningboard.GetOperationLocateState(i)) Then
                Dim strPartNo As String = preactor.ReadFieldString("Orders", "Part No.", i)
                Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", i)

                If (strFirstPartNo = "") Then
                    strFirstPartNo = strPartNo
                End If
                If strBelongsToOrderNo = "PARENT" Then
                    If (Not Information.IsNothing(strPartNo)) Then
                        numOfOrders = numOfOrders + 1
                        Dim dr_int As DataRow = dt.NewRow()
                        dr_int("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", i)
                        dr_int("OrderQuantity") = preactor.ReadFieldString("Orders", "Quantity", i)
                        dr_int("OrderStartDate") = preactor.ReadFieldString("Orders", "Start Time", i)
                        If jobStartDate.ToString = "#1/1/0001 12:00:00 AM#" And Not (preactor.ReadFieldString("Orders", "Start Time", i) = "Unspecified") Then
                            jobStartDate = CDate(preactor.ReadFieldString("Orders", "Start Time", i))
                        Else
                            If CDate(preactor.ReadFieldString("Orders", "Start Time", i)) < jobStartDate Then
                                jobStartDate = CDate(preactor.ReadFieldString("Orders", "Start Time", i))
                            End If
                        End If

                        product = preactor.ReadFieldString("Orders", "Product", i)
                        Dim productRecordNum As Integer = preactor.FindMatchingRecord("Products", "Product", productRecordNum, product)
                        If productRecordNum > 0 Then
                            damage_Percentage = CInt(preactor.ReadFieldString("Products", "Numerical Attribute 3", productRecordNum))
                            If damage_Percentage = 0 Then
                                damage_Percentage = 1
                            End If
                        End If

                        resource = preactor.ReadFieldString("Orders", "Resource", i)

                        If Not ((resource = "Nothing") Or (resource = "Unspecified")) Then
                            Dim resourceRecordNum As Integer = preactor.FindMatchingRecord("Resources", "Name", resourceRecordNum, resource)
                            If resourceRecordNum < 0 Then
                                resource = "Nothing"
                            Else
                                Dim OrderNum As String = preactor.ReadFieldString("Orders", "Order No.", i)
                                orderResourceRate = K201_GetOrderResourceRate(connetionString, OrderNum, resourceRecordNum)

                                line_capacity = CInt(preactor.ReadFieldString("Resources", "K201_Line_Capacity", resourceRecordNum))
                                pitch_Lenght = CDec(preactor.ReadFieldString("Resources", "K201_Pitch_Lenght", resourceRecordNum))
                                line_Speed = CDec(preactor.ReadFieldString("Resources", "K201_Line_Speed", resourceRecordNum))
                            End If
                        End If

                        dt.Rows.Add(dr_int)
                    End If
                End If
            End If
            i = i + 1
        Loop While i <= num

        dt.DefaultView.Sort = "OrderStartDate ASC, OrderNo ASC"

        dt = dt.DefaultView.ToTable()

        Dim dtCal As DataTable = New DataTable()
        Dim dt_s As DataTable = New DataTable()
        ''define variable and assign
        Dim c_CoProductNo_s As DataColumn = New DataColumn("CoProduct", Type.[GetType]("System.String"))
        Dim c_CoProductQuantity_s As DataColumn = New DataColumn("CoProductQuantity", Type.[GetType]("System.String"))
        Dim c_OrderNo_s As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_OrderStartDate_s As DataColumn = New DataColumn("OrderStartDate", Type.[GetType]("System.String"))
        Dim c_CoProductCompleteQuantity_s As DataColumn = New DataColumn("CoProductCompleteQuantity", Type.[GetType]("System.Double"))
        Dim c_CoProductAvilableFormer_s As DataColumn = New DataColumn("AvilableFormer", Type.[GetType]("System.Double"))
        Dim c_CoProduct_SizeCode_s As DataColumn = New DataColumn("CoProductSizeCode", Type.[GetType]("System.String"))
        Dim c_CoProduct_Sequence_s As DataColumn = New DataColumn("CoProductSequence", Type.[GetType]("System.Double"))
        ''Add table column
        dt_s.Columns.Add(c_CoProductNo_s)
        dt_s.Columns.Add(c_CoProductQuantity_s)
        dt_s.Columns.Add(c_OrderNo_s)
        dt_s.Columns.Add(c_OrderStartDate_s)
        dt_s.Columns.Add(c_CoProductAvilableFormer_s)
        dt_s.Columns.Add(c_CoProductCompleteQuantity_s)
        dt_s.Columns.Add(c_CoProduct_SizeCode_s)
        dt_s.Columns.Add(c_CoProduct_Sequence_s)

        Dim totalOrderQty As Double = 0
        For Each product_row As DataRow In dt.Rows
            Dim OrderNo As String = product_row("OrderNo").ToString()
            Dim OrderStartDate As String = product_row("OrderStartDate").ToString()

            Dim coProductRecordNum As Integer = 0

            coProductRecordNum = preactor.FindMatchingRecord("Co-products", "Order No.", coProductRecordNum, OrderNo)
            Dim count As Integer = 1
            While coProductRecordNum > 0
                If (Not Information.IsNothing(preactor.ReadFieldString("Co-products", "Co-product", coProductRecordNum))) Then
                    Dim dt_sr As DataRow = dt_s.NewRow()

                    dt_sr("CoProduct") = preactor.ReadFieldString("Co-products", "Co-product", coProductRecordNum)
                    dt_sr("CoProductQuantity") = preactor.ReadFieldString("Co-products", "Quantity", coProductRecordNum)
                    dt_sr("OrderNo") = OrderNo
                    dt_sr("OrderStartDate") = OrderStartDate
                    dt_sr("CoProductCompleteQuantity") = preactor.ReadFieldString("Co-products", "K201_Completed_Quantity", coProductRecordNum)
                    Dim FormerName As String = preactor.ReadFieldString("Co-products", "K201_Former", coProductRecordNum)
                    Dim avilableFormerCount As Decimal = K201_GetAvilableFormer(connetionString, FormerName)
                    dt_sr("AvilableFormer") = avilableFormerCount
                    dt_sr("CoProductSizeCode") = preactor.ReadFieldString("Co-products", "K201_Size_Code", coProductRecordNum)
                    dt_sr("CoProductSequence") = preactor.ReadFieldString("Co-products", "K201_Size_Sequence", coProductRecordNum)

                    dt_s.Rows.Add(dt_sr)
                    coProductRecordNum = preactor.FindMatchingRecord("Co-products", "Order No.", coProductRecordNum, OrderNo)
                    count = count + 1
                End If
            End While

        Next product_row

        dt_s.DefaultView.Sort = "OrderNo ASC ,CoProductSequence ASC"

        dt_s = dt_s.DefaultView.ToTable()

        ''============================Main Table with  select column============================
        Dim CorProduct_dt As DataTable = dt_s.DefaultView.ToTable(True, "CoProduct")
        Dim order_tb As DataTable = dt_s.DefaultView.ToTable(True, "OrderNo", "CoProduct", "CoProductQuantity", "OrderStartDate", "AvilableFormer")
        ''============================Main Table==========================================
        Dim tbl_formerDetailsHeader As DataTable = New DataTable()
        Dim rowId As DataColumn = New DataColumn("ID", Type.[GetType]("System.Double"))
        tbl_formerDetailsHeader.Columns.Add(rowId)

        Dim former_Size As DataColumn = New DataColumn("#", Type.[GetType]("System.String"))
        tbl_formerDetailsHeader.Columns.Add(former_Size)

        For Each product_row As DataRow In CorProduct_dt.Rows
            Dim coProduct As String = product_row("CoProduct").ToString()
            Dim orderCol As DataColumn = New DataColumn(coProduct, Type.[GetType]("System.String"))
            tbl_formerDetailsHeader.Columns.Add(orderCol)
        Next product_row

        Dim total As DataColumn = New DataColumn("Total", Type.[GetType]("System.Double"))
        tbl_formerDetailsHeader.Columns.Add(total)
        ''Asign row wise row number
        Dim orderdistinct_dt As DataTable = order_tb.DefaultView.ToTable(True, "OrderNo")
        Dim rowId_ As Integer = 10 ''Orders
        For Each product_row As DataRow In orderdistinct_dt.Rows
            Dim co_int As DataRow = tbl_formerDetailsHeader.NewRow()
            co_int("ID") = rowId_
            co_int("#") = product_row("OrderNo").ToString()
            tbl_formerDetailsHeader.Rows.Add(co_int)
            rowId_ = rowId_ + 1
        Next

        rowId_ = 20
        Dim totalcal As DataRow = tbl_formerDetailsHeader.NewRow()
        totalcal("ID") = rowId_
        totalcal("#") = "Total"
        tbl_formerDetailsHeader.Rows.Add(totalcal)

        rowId_ = 25
        Dim totalDamagecal As DataRow = tbl_formerDetailsHeader.NewRow()
        totalDamagecal("ID") = rowId_
        totalDamagecal("#") = "Damage %"
        tbl_formerDetailsHeader.Rows.Add(totalDamagecal)

        rowId_ = 30
        For Each product_row As DataRow In orderdistinct_dt.Rows
            Dim co_int As DataRow = tbl_formerDetailsHeader.NewRow()
            co_int("ID") = rowId_
            co_int("#") = product_row("OrderNo").ToString()
            tbl_formerDetailsHeader.Rows.Add(co_int)
            rowId_ = rowId_ + 1
        Next
        Dim orderCount As Integer = rowId_ - 30

        rowId_ = 41
        Dim totalDamage As DataRow = tbl_formerDetailsHeader.NewRow()
        totalDamage("ID") = rowId_
        totalDamage("#") = "Total With Prov"
        tbl_formerDetailsHeader.Rows.Add(totalDamage)

        rowId_ = 45
        Dim ratioPer As DataRow = tbl_formerDetailsHeader.NewRow()
        ratioPer("ID") = rowId_
        ratioPer("#") = "Total %"
        tbl_formerDetailsHeader.Rows.Add(ratioPer)

        rowId_ = 50
        Dim formerAvailability As DataRow = tbl_formerDetailsHeader.NewRow()
        formerAvailability("ID") = rowId_
        formerAvailability("#") = "Former Availability"
        tbl_formerDetailsHeader.Rows.Add(formerAvailability)

        rowId_ = 55
        Dim former As DataRow = tbl_formerDetailsHeader.NewRow()
        former("ID") = rowId_
        former("#") = "Former %"
        tbl_formerDetailsHeader.Rows.Add(former)


        rowId_ = 60
        Dim formerBalance As DataRow = tbl_formerDetailsHeader.NewRow()
        formerBalance("ID") = rowId_
        formerBalance("#") = "Former Balance"
        tbl_formerDetailsHeader.Rows.Add(formerBalance)

        rowId_ = 65
        Dim lineSpeed As DataRow = tbl_formerDetailsHeader.NewRow()
        lineSpeed("ID") = rowId_
        lineSpeed("#") = "Lines Speed (mts / min)"
        tbl_formerDetailsHeader.Rows.Add(lineSpeed)

        rowId_ = 70
        Dim pitch As DataRow = tbl_formerDetailsHeader.NewRow()
        pitch("ID") = rowId_
        pitch("#") = "Pitch (m)"
        tbl_formerDetailsHeader.Rows.Add(pitch)

        rowId_ = 75
        Dim glovesPerDay As DataRow = tbl_formerDetailsHeader.NewRow()
        glovesPerDay("ID") = rowId_
        glovesPerDay("#") = "No. of Gloves/Day"
        tbl_formerDetailsHeader.Rows.Add(glovesPerDay)

        rowId_ = 80
        Dim pisPerHr As DataRow = tbl_formerDetailsHeader.NewRow()
        pisPerHr("ID") = rowId_
        pisPerHr("#") = "Pieces/Hour"
        tbl_formerDetailsHeader.Rows.Add(pisPerHr)


        rowId_ = 90
        For Each product_row As DataRow In orderdistinct_dt.Rows
            Dim co_int_ As DataRow = tbl_formerDetailsHeader.NewRow()

            co_int_("ID") = rowId_
            co_int_("#") = product_row("OrderNo").ToString()
            tbl_formerDetailsHeader.Rows.Add(co_int_)
            rowId_ = rowId_ + 1

        Next
        ''Assign value to reading row and columns
        For Each formDet As DataRow In tbl_formerDetailsHeader.Rows
            Dim rowValue As String = formDet("#").ToString()
            Dim fdRowId As Integer = CInt(formDet("ID").ToString())
            For Each corpro As DataRow In order_tb.Rows
                Dim order_ As String = corpro("OrderNo").ToString()
                If rowValue = order_ Then
                    Dim sizeColumn As String = corpro("CoProduct").ToString()
                    Dim coProQty As String = corpro("CoProductQuantity").ToString()
                    formDet(sizeColumn) = coProQty
                End If
            Next

            order_tb.DefaultView.Sort = "CoProduct ASC"
            order_tb = order_tb.DefaultView.ToTable()
            If rowValue = "Total" Then
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    Dim sumCorPro As Decimal = 0
                    For Each order As DataRow In order_tb.Rows
                        Dim ordCorProduct As String = order("CoProduct").ToString()
                        If ordCorProduct = corProduct Then
                            sumCorPro = sumCorPro + CDec(order("CoProductQuantity").ToString())
                        End If
                    Next
                    formDet(corProduct) = sumCorPro
                Next
            End If
            If rowValue = "Damage %" Then
                formDet("Total") = damage_Percentage
                damage_Percentage = CDec(((100 - CDec(damage_Percentage)) / 100) / CDec(damage_Percentage))

            End If

            ''Damage % calculation
            If fdRowId >= 30 And fdRowId <= 40 Then
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    Dim orderRow As DataRow = tbl_formerDetailsHeader.Rows(fdRowId - 30)
                    If Not orderRow(corProduct).ToString = "" Then
                        formDet(corProduct) = Math.Round((CDec(orderRow(corProduct)) * damage_Percentage), 2)
                    End If
                Next
            End If
            ''Total Damage
            Dim runningNo As Integer = orderCount + 2
            If fdRowId = 41 Then
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()

                    Dim coTotal As Decimal = 0
                    For i = runningNo To runningNo + orderCount
                        Dim orderRow As DataRow = tbl_formerDetailsHeader.Rows(i)
                        If Not orderRow(corProduct).ToString = "" Then
                            coTotal = coTotal + CDec(orderRow(corProduct))
                        End If
                    Next
                    formDet(corProduct) = Math.Round(coTotal, 2)
                    totalOrderQty = totalOrderQty + coTotal
                Next
                formDet("Total") = Math.Round(totalOrderQty, 2)
            End If
            runningNo = runningNo + orderCount

            '%
            If fdRowId = 45 Then
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()

                    Dim orderRow As DataRow = tbl_formerDetailsHeader.Rows(runningNo)
                    If Not orderRow(corProduct).ToString = "" Then
                        formDet(corProduct) = Math.Round((CDec(orderRow(corProduct)) / totalOrderQty) * 100, 2)
                    End If
                Next
            End If
            runningNo = runningNo + 1


            'Former Availability
            Dim totalFormer As Decimal
            If fdRowId = 50 Then
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    Dim sumCorPro As Decimal = 0
                    For Each order As DataRow In order_tb.Rows
                        Dim ordCorProduct As String = order("CoProduct").ToString()
                        If ordCorProduct = corProduct Then
                            formDet(corProduct) = CDec(order("AvilableFormer").ToString())
                        End If
                    Next
                Next
                Dim orderRow As DataRow = tbl_formerDetailsHeader.Rows(runningNo + 1)
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    totalFormer = totalFormer + CDec(orderRow(corProduct).ToString())
                Next

                ''formDet("Total") = totalFormer
            End If
            runningNo = runningNo + 1

            'Former %
            If fdRowId = 55 Then
                Dim precentage As DataRow = tbl_formerDetailsHeader.Rows(runningNo - 1)
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    formDet(corProduct) = Math.Round(CDec(precentage(corProduct).ToString()) / 100 * line_capacity, 0)
                Next
                formDet("Total") = line_capacity
            End If
            runningNo = runningNo + 1
            'Former Balance
            If fdRowId = 60 Then
                Dim AvbFormer As DataRow = tbl_formerDetailsHeader.Rows(runningNo - 1)
                Dim ReqFormer As DataRow = tbl_formerDetailsHeader.Rows(runningNo)

                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    formDet(corProduct) = Math.Round(CDec(AvbFormer(corProduct).ToString()) - CDec(ReqFormer(corProduct).ToString()), 0)
                Next
            End If
            runningNo = runningNo + 1
            'Lines Speed (mts / min)
            If fdRowId = 65 Then
                formDet("Total") = line_Speed
            End If
            runningNo = runningNo + 1

            'Pitch (m)
            If fdRowId = 70 Then
                formDet("Total") = pitch_Lenght
            End If
            runningNo = runningNo + 1

            'No. of Gloves/Day
            If fdRowId = 75 Then
                Dim FormerPes As DataRow = tbl_formerDetailsHeader.Rows(runningNo - 5) '' Former precentage line 45
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    formDet(corProduct) = Math.Round(CDec(FormerPes(corProduct).ToString()) * orderResourceRate * 24, 0)
                Next
            End If
            runningNo = runningNo + 1

            'Pieces/Hour
            If fdRowId = 80 Then
                Dim FormerPes As DataRow = tbl_formerDetailsHeader.Rows(runningNo - 6) '' Former precentage line 45
                For Each corpro As DataRow In CorProduct_dt.Rows
                    Dim corProduct As String = corpro("CoProduct").ToString()
                    formDet(corProduct) = Math.Round(CDec(FormerPes(corProduct).ToString()) * orderResourceRate, 0)
                Next
            End If
            runningNo = runningNo + 1

            If fdRowId >= 90 And fdRowId <= 100 Then
                If numOfOrders >= fdRowId - 90 Then
                    For Each corpro As DataRow In CorProduct_dt.Rows
                        Dim corProduct As String = corpro("CoProduct").ToString()
                        Dim orderwithWastRow As DataRow = tbl_formerDetailsHeader.Rows(fdRowId - 86)
                        Dim dayRate As DataRow = tbl_formerDetailsHeader.Rows(fdRowId + (numOfOrders - 2) - 76)

                        If Not orderwithWastRow(corProduct).ToString = "" Then
                            If CDec(orderwithWastRow(corProduct)) > 0 And CDec(dayRate(corProduct)) > 0 Then
                                Dim noOfDay As Decimal = 0
                                noOfDay = Math.Round((CDec(orderwithWastRow(corProduct)) / CDec(dayRate(corProduct))), 5)
                                If fdRowId - 90 > 0 Then
                                    Dim previousOrder As DataRow = tbl_formerDetailsHeader.Rows(fdRowId - 76)
                                    If IsDate(previousOrder(corProduct)) Then
                                        Dim orstartDate As DateTime = CDate(previousOrder(corProduct))
                                        formDet(corProduct) = orstartDate.AddDays(noOfDay)
                                    End If
                                Else
                                    formDet(corProduct) = jobStartDate.AddDays(noOfDay)
                                End If
                            End If
                        End If
                    Next
                End If
                numOfOrders = numOfOrders - 1
            End If

        Next

        Dim oForm As New K201_ProductFormarDetails

        Dim tbl_CoProductCompletionDate As DataTable = New DataTable()

        oForm.tblFormerDetailsMain = tbl_formerDetailsHeader
        ''show popup screen
        oForm.Show()
        Return 0
    End Function
#End Region

#Region "Former Ratio Calculation Product"
    ''K201_FormerRatioCalculationProduct this programme will execute when you hit formar ratio calculation button
    Public Function K201_FormerRatioCalculationProduct(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        ''define variable and assign
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim strFirstPartNo As String = ""
        Dim intSumQuanity As Decimal = New Decimal()
        Dim corProducts(preactor.RecordCount("Orders") + 1 - 1, 2) As String
        Dim line_Capacity As Double = 1
        Dim productResourceRate As Double = 1
        Dim line_Type As String = ""
        Dim orderResourceRate As Decimal = 1
        Dim damage_Percentage As Decimal = 1
        Dim jobStartDate As DateTime
        Dim numOfOrders As Integer = 0
        ''Create temp table and column asign
        Dim dtCal As DataTable = New DataTable()
        Dim dt_s As DataTable = New DataTable()
        Dim c_OrderId_s As DataColumn = New DataColumn("OrderId", Type.[GetType]("System.Double"))
        Dim c_ProductNo_s As DataColumn = New DataColumn("Product", Type.[GetType]("System.String"))
        Dim c_ProductQuantity_s As DataColumn = New DataColumn("ProductQuantity", Type.[GetType]("System.String"))
        Dim c_OrderNo_s As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_OrderStartDate_s As DataColumn = New DataColumn("OrderStartDate", Type.[GetType]("System.String"))
        Dim c_ProductCompleteQuantity_s As DataColumn = New DataColumn("ProductCompleteQuantity", Type.[GetType]("System.Double"))
        Dim c_ProductAvilableFormer_s As DataColumn = New DataColumn("AvilableFormer", Type.[GetType]("System.Double"))
        Dim c_Product_SizeCode_s As DataColumn = New DataColumn("ProductSizeCode", Type.[GetType]("System.String"))
        Dim c_Product_Sequence_s As DataColumn = New DataColumn("ProductSequence", Type.[GetType]("System.Double"))
        Dim c_Uom_s As DataColumn = New DataColumn("UOM", Type.[GetType]("System.String"))
        Dim c_FormerType_s As DataColumn = New DataColumn("FormerType", Type.[GetType]("System.String"))
        Dim c_LineType_s As DataColumn = New DataColumn("LineType", Type.[GetType]("System.String"))
        Dim c_LineCapacity_s As DataColumn = New DataColumn("LineCapacity", Type.[GetType]("System.String"))
        Dim c_FormersPerPlatoon_s As DataColumn = New DataColumn("FormersPerPlatoon", Type.[GetType]("System.String"))
        Dim c_OrderResourceRate_s As DataColumn = New DataColumn("OrderResourceRate", Type.[GetType]("System.String"))
        Dim c_SecondaryConstraintQty_s As DataColumn = New DataColumn("SecondaryConstraintQty", Type.[GetType]("System.Double"))
        Dim c_ProductTotalAvilableFormer_s As DataColumn = New DataColumn("TotalAvilableFormer", Type.[GetType]("System.Double"))


        dt_s.Columns.Add(c_OrderId_s)
        dt_s.Columns.Add(c_ProductNo_s)
        dt_s.Columns.Add(c_ProductQuantity_s)
        dt_s.Columns.Add(c_OrderNo_s)
        dt_s.Columns.Add(c_OrderStartDate_s)
        dt_s.Columns.Add(c_ProductAvilableFormer_s)
        dt_s.Columns.Add(c_ProductCompleteQuantity_s)
        dt_s.Columns.Add(c_Product_SizeCode_s)
        dt_s.Columns.Add(c_Product_Sequence_s)
        dt_s.Columns.Add(c_Uom_s)
        dt_s.Columns.Add(c_FormerType_s)
        dt_s.Columns.Add(c_LineType_s)
        dt_s.Columns.Add(c_LineCapacity_s)
        dt_s.Columns.Add(c_FormersPerPlatoon_s)
        dt_s.Columns.Add(c_OrderResourceRate_s)
        dt_s.Columns.Add(c_SecondaryConstraintQty_s)
        dt_s.Columns.Add(c_ProductTotalAvilableFormer_s)

        Dim totalOrderQty As Double = 0
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim FirstSelectedProduct As String
        Dim FirstSelectedResource As String
        Dim isValidateOrder As Integer = 0
        Dim isValidProduct As Integer = 0
        Dim totalGlovePairQty As Double = 0


        Dim y As Integer = 1
        Dim firstRecord As Integer = 1
        ''loop through order table if block selected programe will  execut  and calculat the ratio
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                ''Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", y)
                ''If strBelongsToOrderNo = "PARENT" Then
                Dim intOperationNumber As Integer = preactor.ReadFieldInt("Orders", "Op. No.", y)
                If intOperationNumber = 10 Then
                    Dim orno As String = preactor.ReadFieldString("Orders", "Order No.", y)
                    If firstRecord = 1 Then
                        FirstSelectedProduct = preactor.ReadFieldString("Orders", "String Attribute 4", y)
                        FirstSelectedResource = preactor.ReadFieldString("Orders", "Resource", y)
                        firstRecord = firstRecord + 1
                    End If
                    If Not (FirstSelectedProduct = preactor.ReadFieldString("Orders", "String Attribute 4", y)) Then
                        isValidProduct = -1
                    End If
                    If Not (FirstSelectedResource = preactor.ReadFieldString("Orders", "Resource", y)) Then
                        isValidateOrder = -1
                    End If
                    ''End If
                End If
            End If
            y = y + 1
        Loop While y <= num
        ''check same product and resorces are same if not message will popup
        If isValidateOrder = -1 And isValidProduct = -1 Then
            MsgBox("Former ratio calculation cannot countinue, Selected orders have different Product OR Resources", vbCritical, "")
        ElseIf isValidateOrder = -1 Then
            MsgBox("Former ratio calculation cannot countinue, Selected orders have different Resources", vbCritical, "")
        ElseIf isValidProduct = -1 Then
            MsgBox("Former ratio calculation cannot countinue, Selected orders have different Product", vbCritical, "")
        Else
            Dim i As Integer = 1
            Do
                If (planningboard.GetOperationLocateState(i)) Then

                    ''Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", i)
                    ''If strBelongsToOrderNo = "PARENT" Then
                    Dim intOperationNumber As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
                    If intOperationNumber = 10 Then
                        Dim dt_sr As DataRow = dt_s.NewRow()
                        Dim orId As Integer = preactor.ReadFieldInt("Orders", "Number", i)

                        dt_sr("OrderId") = orId

                        dt_sr("Product") = preactor.ReadFieldString("Orders", "K201 Size Code", i)
                        dt_sr("ProductQuantity") = preactor.ReadFieldString("Orders", "Quantity", i)

                        dt_sr("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", i)
                        If jobStartDate.ToString = "#1/1/0001 12:00:00 AM#" And Not (preactor.ReadFieldString("Orders", "Start Time", i) = "Unspecified") Then
                            jobStartDate = CDate(preactor.ReadFieldString("Orders", "Start Time", i))
                        Else
                            If Not preactor.ReadFieldString("Orders", "Start Time", i) = "Unspecified" Then
                                If CDate(preactor.ReadFieldString("Orders", "Start Time", i)) < jobStartDate Then
                                    jobStartDate = CDate(preactor.ReadFieldString("Orders", "Start Time", i))
                                End If
                            End If
                        End If

                        dt_sr("OrderStartDate") = jobStartDate

                        dt_sr("ProductCompleteQuantity") = preactor.ReadFieldString("Orders", "K201 Completed Quantity", 1)
                        Dim FormerType As String = preactor.ReadFieldString("Orders", "Table Attribute 1", i)
                        dt_sr("FormerType") = FormerType

                        Dim resource As String
                        resource = preactor.ReadFieldString("Orders", "Resource", i)
                        If Not ((resource = "Nothing") Or (resource = "Unspecified")) Then
                            Dim resourceRecordNum As Integer = preactor.FindMatchingRecord("Resources", "Name", resourceRecordNum, resource)
                            If resourceRecordNum < 0 Then
                                resource = "Nothing"
                            Else
                                Dim orerderNo As String = preactor.ReadFieldString("Orders", "Order No.", i)
                                line_Capacity = CDec(preactor.ReadFieldDouble("Resources", "K201 Line Capacity", resourceRecordNum))
                                line_Type = (preactor.ReadFieldString("Resources", "K201 Line Type", resourceRecordNum))
                                dt_sr("LineCapacity") = line_Capacity
                                dt_sr("LineType") = line_Type
                            End If
                        End If

                        Dim avilableFormerCount As Decimal = K201_GetAvilableFormer(connetionString, preactor.ReadFieldString("Orders", "Order No.", i))
                        dt_sr("AvilableFormer") = avilableFormerCount

                        Dim totalAvilableFormerCount As Decimal = K201_GetTotalAvilableFormer(connetionString, preactor.ReadFieldString("Orders", "Order No.", i))
                        dt_sr("TotalAvilableFormer") = totalAvilableFormerCount

                        productResourceRate = K201_GetProductResourceRate(connetionString, preactor.ReadFieldString("Orders", "Order No.", i))
                        orderResourceRate = K201_GetOrderResourceRate(connetionString, preactor.ReadFieldString("Orders", "Order No.", i))
                        dt_sr("OrderResourceRate") = orderResourceRate

                        If (line_Type = "P") Then
                            Dim formersPerPlatoon As Decimal = K201_GetFormersPerPlatoon(connetionString, preactor.ReadFieldString("Orders", "Order No.", i))
                            dt_sr("FormersPerPlatoon") = formersPerPlatoon
                        Else
                            dt_sr("FormersPerPlatoon") = 0
                        End If

                        dt_sr("ProductSizeCode") = preactor.ReadFieldString("Orders", "K201 Size Code", i)
                        dt_sr("ProductSequence") = preactor.ReadFieldString("Orders", "K201 Size Sequance", i)
                        dt_sr("UOM") = preactor.ReadFieldString("Orders", "K201 Uom", i)
                        dt_sr("SecondaryConstraintQty") = preactor.ReadFieldString("Orders", "Numerical Attribute 5", i)


                        Dim productRecordNum As Integer = preactor.FindMatchingRecord("Products", "Product", productRecordNum, preactor.ReadFieldString("Orders", "Product", i))
                        If productRecordNum > 0 Then
                            damage_Percentage = CDec(preactor.ReadFieldString("Products", "Numerical Attribute 3", productRecordNum))
                            If damage_Percentage = 0 Then
                                damage_Percentage = 1
                            End If
                        End If

                        dt_s.Rows.Add(dt_sr)

                    End If
                End If

                i = i + 1
            Loop While i <= num

            dt_s.DefaultView.Sort = "OrderNo ASC ,ProductSequence ASC"

            dt_s = dt_s.DefaultView.ToTable()


            ''============================Main Table============================
            Dim order_tb As DataTable = dt_s.DefaultView.ToTable(True, "OrderNo", "Product", "ProductQuantity", "OrderStartDate", "AvilableFormer", "FormersPerPlatoon", "OrderResourceRate", "OrderId", "SecondaryConstraintQty", "TotalAvilableFormer")


            Dim Product_dts As DataTable = dt_s.DefaultView.ToTable(True, "Product", "ProductSequence")
            Product_dts.DefaultView.Sort = "ProductSequence ASC"
            Product_dts = Product_dts.DefaultView.ToTable()
            Dim Product_dt As DataTable = Product_dts.DefaultView.ToTable(True, "Product")



            Dim tb_orders As DataTable = dt_s.DefaultView.ToTable(True, "OrderId")
            Dim str_orders As String = ""
            For Each ord_ As DataRow In tb_orders.Rows
                If str_orders = "" Then
                    str_orders = str_orders + ord_("OrderId").ToString()
                Else
                    str_orders = str_orders + "|" + ord_("OrderId").ToString()
                End If
            Next
            ''Get extra formar form that othe plant
            Dim extra_formers As DataTable
            extra_formers = K201_GetExtraFormerTbl(connetionString, str_orders)

            ''======================================================================
            Dim tbl_formerDetailsHeader As DataTable = New DataTable()
            Dim rowId As DataColumn = New DataColumn("ID", Type.[GetType]("System.Double"))
            tbl_formerDetailsHeader.Columns.Add(rowId)

            Dim former_Size As DataColumn = New DataColumn("#", Type.[GetType]("System.String"))
            tbl_formerDetailsHeader.Columns.Add(former_Size)

            '####################################################

            Dim OrderIdNew As DataColumn = New DataColumn("OrderID", Type.[GetType]("System.String"))
            tbl_formerDetailsHeader.Columns.Add(OrderIdNew)

            For Each product_row As DataRow In Product_dt.Rows
                Dim coProduct As String = product_row("Product").ToString()
                Dim orderCol As DataColumn = New DataColumn(coProduct, Type.[GetType]("System.String"))
                tbl_formerDetailsHeader.Columns.Add(orderCol)
            Next product_row

            Dim total As DataColumn = New DataColumn("Total", Type.[GetType]("System.String"))
            tbl_formerDetailsHeader.Columns.Add(total)
            Dim columnsArr() As String = {"OrderNo", "OrderId"}
            Dim orderdistinct_dt As DataTable = order_tb.DefaultView.ToTable(True, columnsArr)
            ''Assign temp table row and column according to  the selected product sizes 
            Dim rowId_ As Integer = 10 ''Orders will  filling 10  to  20 number of orders

            For Each product_row As DataRow In orderdistinct_dt.Rows
                Dim co_int As DataRow = tbl_formerDetailsHeader.NewRow()

                co_int("ID") = rowId_
                co_int("#") = product_row("OrderNo").ToString()
                co_int("OrderID") = product_row("OrderId").ToString()
                tbl_formerDetailsHeader.Rows.Add(co_int)
                rowId_ = rowId_ + 1
            Next

            rowId_ = 30
            Dim totalcal As DataRow = tbl_formerDetailsHeader.NewRow()
            totalcal("ID") = rowId_
            totalcal("#") = "Total"
            tbl_formerDetailsHeader.Rows.Add(totalcal)


            rowId_ = 40
            Dim ratioPer As DataRow = tbl_formerDetailsHeader.NewRow()
            ratioPer("ID") = rowId_
            ratioPer("#") = "Total %"
            tbl_formerDetailsHeader.Rows.Add(ratioPer)

            rowId_ = 50
            Dim formerAvailability As DataRow = tbl_formerDetailsHeader.NewRow()
            formerAvailability("ID") = rowId_
            If line_Type = "C" Then
                formerAvailability("#") = "Plant Former Availability"
            ElseIf line_Type = "P" Then
                formerAvailability("#") = "Plant Platoon Availability"
            Else
                formerAvailability("#") = "Not defined"
            End If
            tbl_formerDetailsHeader.Rows.Add(formerAvailability)

            rowId_ = 55
            Dim totalFormerAvailability As DataRow = tbl_formerDetailsHeader.NewRow()
            totalFormerAvailability("ID") = rowId_
            If line_Type = "C" Then
                totalFormerAvailability("#") = "Total Former Availability"
            ElseIf line_Type = "P" Then
                totalFormerAvailability("#") = "Total Platoon Availability"
            Else
                totalFormerAvailability("#") = "Not defined"
            End If
            tbl_formerDetailsHeader.Rows.Add(totalFormerAvailability)

            rowId_ = 60
            Dim lineSpeed As DataRow = tbl_formerDetailsHeader.NewRow()
            lineSpeed("ID") = rowId_
            lineSpeed("#") = "Lines Capacity"
            tbl_formerDetailsHeader.Rows.Add(lineSpeed)

            rowId_ = 105
            Dim prr As DataRow = tbl_formerDetailsHeader.NewRow()
            prr("ID") = rowId_
            prr("#") = "Product Resource Rate"
            tbl_formerDetailsHeader.Rows.Add(prr)


            rowId_ = 110
            Dim former As DataRow = tbl_formerDetailsHeader.NewRow()
            former("ID") = rowId_
            If line_Type = "C" Then
                former("#") = "Former Requirement"
            ElseIf line_Type = "P" Then
                former("#") = "Platoon Ratio"
            Else
                former("#") = "Not defined"
            End If
            tbl_formerDetailsHeader.Rows.Add(former)

            rowId_ = 115
            Dim formerEntered As DataRow = tbl_formerDetailsHeader.NewRow()
            formerEntered("ID") = rowId_
            If line_Type = "C" Then
                formerEntered("#") = "Enter Former"
            ElseIf line_Type = "P" Then
                formerEntered("#") = "Allocated Platoons"
            Else
                formerEntered("#") = "Not defined"
            End If
            tbl_formerDetailsHeader.Rows.Add(formerEntered)


            rowId_ = 120
            Dim formerBalance As DataRow = tbl_formerDetailsHeader.NewRow()
            formerBalance("ID") = rowId_
            If line_Type = "C" Then
                formerBalance("#") = "Former Balance"
            ElseIf line_Type = "P" Then
                formerBalance("#") = "Platoon Balance"
            Else
                formerBalance("#") = "Not defined"
            End If
            tbl_formerDetailsHeader.Rows.Add(formerBalance)

            rowId_ = 125
            Dim orr As DataRow = tbl_formerDetailsHeader.NewRow()
            orr("ID") = rowId_
            orr("#") = "Order Resource Rate"
            tbl_formerDetailsHeader.Rows.Add(orr)


            ''==============If Platoon only this will execute=================
            If line_Type = "P" Then
                rowId_ = 130
                Dim rowfoprpl As DataRow = tbl_formerDetailsHeader.NewRow()
                rowfoprpl("ID") = rowId_
                rowfoprpl("#") = "Formers Per Platoon"
                tbl_formerDetailsHeader.Rows.Add(rowfoprpl)

                'rowId_ = 131
                'Dim glproducedQty As DataRow = tbl_formerDetailsHeader.NewRow()
                'glproducedQty("ID") = rowId_
                'glproducedQty("#") = "Total Gloves"
                'tbl_formerDetailsHeader.Rows.Add(glproducedQty)

                rowId_ = 132
                Dim glpair As DataRow = tbl_formerDetailsHeader.NewRow()
                glpair("ID") = rowId_
                glpair("#") = "Produced Gloves In Pairs"
                tbl_formerDetailsHeader.Rows.Add(glpair)

                rowId_ = 135
                Dim glovrat As DataRow = tbl_formerDetailsHeader.NewRow()
                glovrat("id") = rowId_
                glovrat("#") = "Produced size ratio"
                tbl_formerDetailsHeader.Rows.Add(glovrat)


                rowId_ = 136
                Dim glovDif As DataRow = tbl_formerDetailsHeader.NewRow()
                glovDif("ID") = rowId_
                glovDif("#") = "Produced Glove Ratio Difference"
                tbl_formerDetailsHeader.Rows.Add(glovDif)

                'rowId_ = 140
                'Dim apq As DataRow = tbl_formerDetailsHeader.NewRow()
                'apq("ID") = rowId_
                'apq("#") = "Adjusted Platoon Qty"
                'tbl_formerDetailsHeader.Rows.Add(apq)

                'rowId_ = 145
                'Dim agrd As DataRow = tbl_formerDetailsHeader.NewRow()
                'agrd("ID") = rowId_
                'agrd("#") = "Adjusted Glove Pairs."
                'tbl_formerDetailsHeader.Rows.Add(agrd)

                'rowId_ = 150
                'Dim agrdra As DataRow = tbl_formerDetailsHeader.NewRow()
                'agrdra("ID") = rowId_
                'agrdra("#") = "Adjusted Glove Ratio"
                'tbl_formerDetailsHeader.Rows.Add(agrdra)

                'rowId_ = 155
                'Dim agglowdiff As DataRow = tbl_formerDetailsHeader.NewRow()
                'agglowdiff("ID") = rowId_
                'agglowdiff("#") = "Adjusted Glove Ratio Diff."
                'tbl_formerDetailsHeader.Rows.Add(agglowdiff)

            End If

            ''============================Data Filling ===============================================================================
            Dim maxRate As Decimal = 0
            Dim minRate As Decimal = 0
            For Each formDet As DataRow In tbl_formerDetailsHeader.Rows
                Dim rowValue As String = formDet("#").ToString()
                Dim fdRowId As Integer = CInt(formDet("ID").ToString())

                Dim qtyRowValue As String = formDet("OrderID").ToString()

                For Each corpro As DataRow In order_tb.Rows
                    Dim order_ As String = corpro("OrderId").ToString()
                    If qtyRowValue = order_ Then
                        Dim sizeColumn As String = corpro("Product").ToString()
                        Dim coProQty As Decimal = CDec(corpro("ProductQuantity").ToString())
                        formDet(sizeColumn) = coProQty.ToString("N0")
                    End If
                Next

                order_tb.DefaultView.Sort = "Product ASC"
                order_tb = order_tb.DefaultView.ToTable()

                ''30
                If rowValue = "Total" Then
                    Dim totalOrder As Decimal = 0
                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()
                        Dim sumCorPro As Decimal = 0
                        For Each order As DataRow In order_tb.Rows
                            Dim ordCorProduct As String = order("Product").ToString()
                            If ordCorProduct = corProduct Then
                                sumCorPro = sumCorPro + CDec(order("ProductQuantity").ToString())
                            End If
                        Next
                        formDet(corProduct) = sumCorPro.ToString("N0")
                        totalOrder = totalOrder + sumCorPro
                    Next
                    formDet("Total") = totalOrder.ToString("N0")
                    totalOrderQty = Math.Round(totalOrder, 0)
                End If

                '%
                If fdRowId = 40 Then
                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()

                        Dim orderRow As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count))
                        If Not orderRow(corProduct).ToString = "" Then
                            formDet(corProduct) = Math.Round((CDec(orderRow(corProduct)) / totalOrderQty) * 100, 2)
                        End If
                    Next
                    formDet("Total") = 100
                End If

                'To get plant Former Availability
                Dim availableFormer As Decimal
                If fdRowId = 50 Then
                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()
                        Dim sumCorPro As Decimal = 0
                        For Each order As DataRow In order_tb.Rows
                            Dim ordCorProduct As String = order("Product").ToString()
                            If ordCorProduct = corProduct Then
                                formDet(corProduct) = CDec(order("AvilableFormer").ToString()).ToString("N0")
                                availableFormer = availableFormer + CDec(order("AvilableFormer").ToString())
                            End If
                        Next
                    Next
                End If

                'To get total Former Availability
                Dim totalAvailableFormer As Decimal
                If fdRowId = 55 Then
                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()
                        Dim sumCorPro As Decimal = 0
                        For Each order As DataRow In order_tb.Rows
                            Dim ordCorProduct As String = order("Product").ToString()
                            If ordCorProduct = corProduct Then
                                formDet(corProduct) = CDec(order("TotalAvilableFormer").ToString()).ToString("N0")
                                totalAvailableFormer = totalAvailableFormer + CDec(order("TotalAvilableFormer").ToString())
                            End If
                        Next
                    Next
                End If

                If fdRowId = 60 Then
                    formDet("Total") = line_Capacity.ToString("N0")
                End If

                If fdRowId = 105 Then
                    formDet("Total") = productResourceRate.ToString("N0")
                End If
                'Former %
                If fdRowId = 110 Then
                    Dim precentage As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 1)
                    Dim totalPrecentage As Decimal
                    For Each corpro As DataRow In Product_dt.Rows
                            Dim corProduct As String = corpro("Product").ToString()
                            formDet(corProduct) = (CDec(precentage(corProduct).ToString()) / 100 * line_Capacity).ToString("N0")
                            totalPrecentage = totalPrecentage + CDec(Math.Round(CDec(precentage(corProduct).ToString()) / 100 * line_Capacity, 0))
                        Next
                    formDet("Total") = totalPrecentage.ToString("N0")
                End If

                ''===Get the default former required quantity into the Enter former field when calculate the former ratio calculation 
                'for the first time

                'Former Manual Values

                'If fdRowId = 115 Then
                '    'Dim precentage As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 1)
                '    Dim totalFormersAssign As Decimal
                '    For Each order As DataRow In order_tb.Rows
                '        Dim corProduct As String = order("Product").ToString()
                '        formDet(corProduct) = CDec(order("SecondaryConstraintQty").ToString()).ToString("N0")
                '        totalFormersAssign = totalFormersAssign + CInt(formDet(corProduct))
                '        'totalFormersAssign = totalFormersAssign + CDec(order("SecondaryConstraintQty").ToString())
                '    Next
                '    formDet("Total") = totalFormersAssign.ToString("N0")


                If fdRowId = 115 Then
                    Dim totalFormersAssign As Decimal
                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()
                        For Each order As DataRow In order_tb.Rows
                            Dim ordCorProduct As String = order("Product").ToString()
                            If ordCorProduct = corProduct Then
                                formDet(corProduct) = CDec(order("SecondaryConstraintQty").ToString()).ToString("N0")
                                totalFormersAssign = totalFormersAssign + CDec(order("SecondaryConstraintQty").ToString())
                            End If
                        Next
                    Next
                    'formDet("Total") = totalFormersAssign.ToString("N0")

                    If totalFormersAssign = 0 Then
                        Dim precentage As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 1)
                        Dim totalPrecentage As Decimal
                        For Each corpro As DataRow In Product_dt.Rows
                            Dim corProduct As String = corpro("Product").ToString()
                            formDet(corProduct) = (CDec(precentage(corProduct).ToString()) / 100 * line_Capacity).ToString("N0")
                            totalPrecentage = totalPrecentage + CDec(Math.Round(CDec(precentage(corProduct).ToString()) / 100 * line_Capacity, 0))
                        Next
                        'formDet("Total") = totalPrecentage.ToString("N0")
                    End If

                End If

                'Former Balance
                If fdRowId = 120 Then
                    Dim AvbFormer As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 2)
                    Dim ReqFormer As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 6)

                    For Each corpro As DataRow In Product_dt.Rows
                        Dim corProduct As String = corpro("Product").ToString()
                        formDet(corProduct) = (CDec(AvbFormer(corProduct).ToString()) - CDec(ReqFormer(corProduct).ToString())).ToString("N0")
                    Next
                End If
                ''Order Resources rate
                If fdRowId = 125 Then
                    Dim precentage As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 1)

                    For Each pro As DataRow In Product_dt.Rows
                        Dim sProduct As String = pro("Product").ToString()
                        Dim sumCorPro As Decimal = 0
                        For Each order As DataRow In order_tb.Rows
                            Dim product As String = order("Product").ToString()
                            If product = sProduct Then
                                formDet(sProduct) = (CDec(precentage(sProduct).ToString()) / 100 * productResourceRate).ToString("N2")
                            End If
                        Next
                    Next
                End If

                'FormersPerPlatoon
                If line_Type = "P" Then
                    If fdRowId = 130 Then
                        For Each corpro As DataRow In Product_dt.Rows
                            Dim corProduct As String = corpro("Product").ToString()
                            For Each order As DataRow In order_tb.Rows
                                Dim ordCorProduct As String = order("Product").ToString()
                                If ordCorProduct = corProduct Then
                                    formDet(corProduct) = order("FormersPerPlatoon").ToString()
                                End If
                            Next
                        Next
                    End If

                    Dim totalGlovPai As Decimal = 0
                    If fdRowId = 132 Then
                        Dim formersPerPlatoon As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 10)
                        Dim platoonRequrment As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 7)

                        For Each corpro As DataRow In Product_dt.Rows
                            Dim corProduct As String = corpro("Product").ToString()
                            formDet(corProduct) = Math.Round(((CDec(platoonRequrment(corProduct).ToString()) * CDec(formersPerPlatoon(corProduct).ToString())) / 2), 1)
                            totalGlovPai = totalGlovPai + Math.Round(((CDec(platoonRequrment(corProduct).ToString()) * CDec(formersPerPlatoon(corProduct).ToString())) / 2), 1)
                        Next
                        formDet("Total") = totalGlovPai
                        totalGlovePairQty = Math.Round(totalGlovPai, 0)
                    End If

                    If fdRowId = 135 Then
                        Dim glovPair As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 11)

                        For Each corpro As DataRow In Product_dt.Rows
                            Dim corProduct As String = corpro("Product").ToString()
                            formDet(corProduct) = Math.Round((CDec(glovPair(corProduct).ToString()) / totalGlovePairQty) * 100, 2)
                            'If Not CDec(glovPair("Total").ToString()) = 0 Then
                            'If Not CDec(Replace(glovPair("Total").ToString().Trim(CChar(",")), ".", "")) = 0 Then
                            '    formDet(corProduct) = Math.Round(((CDec(glovPair(corProduct).ToString()) * 100) / CDec(glovPair("Total").ToString())), 2)
                            'Else
                            '    MsgBox("TotalGlow Calculation error",, "Error")
                            'End If
                        Next
                        formDet("Total") = 100
                    End If

                    Try
                        ''Glove Ratio Diff.
                        If fdRowId = 136 Then
                            ''Total %
                            Dim TotalPer As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 1)
                            ''Produced Size ratio
                            Dim glovRate As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 12)

                            For Each corpro As DataRow In Product_dt.Rows
                                Dim corProduct As String = corpro("Product").ToString()
                                formDet(corProduct) = Math.Round(CDec(TotalPer(corProduct).ToString()) - CDec(glovRate(corProduct).ToString()), 2)
                            Next
                        End If

                    Catch ex As Exception
                        MsgBox("Formers Per Platoon is Not Given", vbOKCancel, "Platoon Ratio Calculation Error")

                    End Try

                    'Commented ForTesting
                    ''Adjusted Platoon Qty
                    'If fdRowId = 140 Then
                    '    ''Total %
                    '    Dim platoonP As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 6)
                    '    For Each corpro As DataRow In Product_dt.Rows
                    '        Dim product As String = corpro("Product").ToString()
                    '        formDet(product) = platoonP(product).ToString()
                    '    Next
                    '    Dim glratdiff As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 12)
                    '    Dim counter As Integer = 0
                    '    For Each corpro As DataRow In order_tb.Rows
                    '        Dim corproduct As String = corpro("formersperplatoon").ToString()
                    '        If counter = 0 Then
                    '            maxRate = CDec(glratdiff(corproduct).ToString)
                    '            minRate = CDec(glratdiff(corproduct).ToString)
                    '        Else
                    '            If (maxRate < CDec(glratdiff(corproduct).ToString)) Then
                    '                maxRate = CDec(glratdiff(corproduct).ToString)
                    '            End If
                    '            If (minRate > CDec(glratdiff(corproduct).ToString)) Then
                    '                minRate = CDec(glratdiff(corproduct).ToString)
                    '            End If
                    '        End If
                    '        counter = counter + 1
                    '    Next
                    '    ''Add Most minimum value to +1 
                    '    Dim glratdiff_ As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 6)
                    '    Dim glratdiffRate As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 12)

                    '    For Each corpro As DataRow In Product_dt.Rows
                    '        Dim product_ As String = corpro("Product").ToString()
                    '        If CDec(glratdiffRate(product_).ToString()) = minRate Then
                    '            formDet(product_) = CDec(glratdiff_(product_).ToString()) + 1
                    '        End If
                    '        If CDec(glratdiffRate(product_).ToString()) = maxRate Then
                    '            formDet(product_) = CDec(glratdiff_(product_).ToString()) - 1
                    '        End If
                    '    Next

                    'End If

                    '''Adjusted Glove Pairs.
                    'If fdRowId = 145 Then

                    '    Dim totalGlovPaiCal As Decimal = 0
                    '    Dim formersPerPlatoon As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 9)
                    '    Dim platoonRequrment As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 13)

                    '    For Each corpro As DataRow In Product_dt.Rows
                    '        Dim corProduct As String = corpro("Product").ToString()
                    '        formDet(corProduct) = Math.Round(((CDec(platoonRequrment(corProduct).ToString()) * CDec(formersPerPlatoon(corProduct).ToString())) / 2), 1)
                    '        totalGlovPaiCal = totalGlovPaiCal + Math.Round(((CDec(platoonRequrment(corProduct).ToString()) * CDec(formersPerPlatoon(corProduct).ToString())) / 2), 1)
                    '    Next
                    '    formDet("Total") = totalGlovPaiCal
                    'End If
                    '''Adjusted Glove Ratio
                    'If fdRowId = 150 Then
                    '    Dim glovPair As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 10)

                    '    For Each corpro As DataRow In Product_dt.Rows
                    '        Dim corProduct As String = corpro("Product").ToString()
                    '        formDet(corProduct) = Math.Round(((CDec(glovPair(corProduct).ToString()) * 100) / CDec(glovPair("Total").ToString())), 2)
                    '    Next
                    'End If
                    '''Adjusted Glove Ratio Diff.
                    'If fdRowId = 155 Then
                    '    ''Adjusted Total %
                    '    Dim TotalPer As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 14)
                    '    ''Adjusted Glove Ratio
                    '    Dim glovRate As DataRow = tbl_formerDetailsHeader.Rows(CInt(order_tb.Rows.Count) + 3)
                    '    For Each corpro As DataRow In Product_dt.Rows
                    '        Dim corProduct As String = corpro("Product").ToString()
                    '        formDet(corProduct) = Math.Round(CDec(TotalPer(corProduct).ToString()) - CDec(glovRate(corProduct).ToString()), 2)

                    '    Next
                    'End If
                End If
            Next
            '' create product former details window declaration and popup show
            Dim oForm As New K201_ProductFormarDetails()
            Dim K201_SizeDetails As DataTable = New DataTable()
            Dim K201_OrderResourceRate As DataTable
            oForm.tblFormerDetailsMain = tbl_formerDetailsHeader
            oForm.tblSize = Product_dt
            oForm.tblOrder = orderdistinct_dt
            oForm.connetionString = connetionString
            oForm.lineType = line_Type
            oForm.ShowDialog()
            Dim resRowId As Integer = 0
            Try
                If Not oForm.tbltblOrderRate_gl Is Nothing Then

                    If oForm.tbltblOrderRate_gl.Columns.Count > 0 Then
                        K201_OrderResourceRate = oForm.tbltblOrderRate_gl
                        Dim orderId As String = ""
                        Dim orderNo As String = ""
                        Dim rate As Decimal = 0
                        Dim formarRatio As Decimal

                        Dim orderResourceTbl As DataTable = New DataTable()
                        Dim order_ss As DataColumn = New DataColumn("Order", Type.[GetType]("System.String"))
                        Dim resource_ss As DataColumn = New DataColumn("Resource", Type.[GetType]("System.String"))
                        Dim resourceId_ss As DataColumn = New DataColumn("ResourceRowId", Type.[GetType]("System.String"))
                        Dim resourceStartTime_ss As DataColumn = New DataColumn("ResourceStartTime", Type.[GetType]("System.String"))
                        Dim resourceEndTime_ss As DataColumn = New DataColumn("ResourceEndTime", Type.[GetType]("System.String"))
                        Dim state_ss As DataColumn = New DataColumn("State", Type.[GetType]("System.String"))

                        orderResourceTbl.Columns.Add(order_ss)
                        orderResourceTbl.Columns.Add(resource_ss)
                        orderResourceTbl.Columns.Add(resourceId_ss)
                        orderResourceTbl.Columns.Add(resourceStartTime_ss)
                        orderResourceTbl.Columns.Add(resourceEndTime_ss)
                        orderResourceTbl.Columns.Add(state_ss)

                        Dim ResourceRowId As Integer
                        Dim ResourceStartTime As DateTime
                        Dim ResourceEndTime As DateTime
                        'Dim orderRecordNum_ As Integer
                        'orderRecordNum_ = preactor.FindMatchingRecord("Orders", "Order No.", orderRecordNum_, "DC201902702.1-YELLOW-ARO-L")

                        If Not ((K201_OrderResourceRate Is DBNull.Value) Or (K201_OrderResourceRate Is Nothing)) Then
                            Dim blockSelectionId As Integer = CInt(K201_GenarateNewBlockSelectionIdFormer(connetionString))

                            For Each orderr As DataRow In orderdistinct_dt.Rows
                                orderNo = orderr("OrderNo").ToString()
                                Dim order_Num As Integer = 0
                                order_Num = preactor.FindMatchingRecord("Orders", "Order No.", order_Num, orderNo)
                                Dim resourceName As String = preactor.ReadFieldString("Orders", "Resource", order_Num)

                                resRowId = preactor.FindMatchingRecord("Resources", "Name", resRowId, resourceName)

                                If resRowId > 0 Then

                                    Dim ors As DataRow = orderResourceTbl.NewRow()
                                    ors("Order") = order_Num
                                    ors("Resource") = resourceName
                                    ors("ResourceRowId") = resRowId
                                    ors("State") = "A"
                                    orderResourceTbl.Rows.Add(ors)
                                End If

                            Next
                            resRowId = 0
                            Dim orderRecordNum As Integer = 0
                            'Dim parentOrderRecordNum As Integer = 0
                            Dim orderIdRecordNum As Integer = 0
                            Dim orderRcNum As Integer = 0

                            ''===================== To calculate the  former ratio========
                            Dim totalAssignedFormerQty As Decimal = 0
                            'For Each ordersrow As DataRow In K201_OrderResourceRate.Rows
                            '    Dim AssignedFormerQty As Decimal = CDec(ordersrow("FormarRatio").ToString())
                            '    totalAssignedFormerQty = totalAssignedFormerQty + AssignedFormerQty
                            'Next
                            For Each ordersrow As DataRow In K201_OrderResourceRate.Rows
                                totalAssignedFormerQty = CDec(ordersrow("TotalFormers").ToString())
                            Next

                            For Each orderr As DataRow In K201_OrderResourceRate.Rows
                                orderId = orderr("OrderID").ToString()
                                orderNo = orderr("OrderNum").ToString()
                                rate = CDec(orderr("Rate").ToString())
                                formarRatio = CDec(orderr("FormarRatio").ToString())

                                'orderRecordNum = preactor.FindMatchingRecord("Orders", "Number", orderRecordNum, orderId)
                                'MsgBox(orderRecordNum)
                                Dim k As Integer = 1
                                'MsgBox(orderId)

                                Do
                                    If CInt(orderId) = preactor.ReadFieldInt("Orders", "Number", k) Then
                                        orderRecordNum = k
                                        'MsgBox(orderRecordNum)
                                    End If
                                    k = k + 1
                                Loop While k <= num
                                'Dim n As Integer

                                'For i = 0 To 4
                                '    orderRecordNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRecordNum, orderNo)
                                '    'Dim a As New PrimaryKey(CInt(orderId))
                                '    'orderRecordNum = preactor.GetRecordNumber("Orders", a)
                                '    '' orderRecordNum = preactor.ReadFieldInt("Orders", "Number", CInt(orderId) - 1)

                                '    'parentOrderRecordNum = preactor.FindMatchingRecord("Orders", "Belongs to Order No.", parentOrderRecordNum, orderRecordNum)

                                '    If ("Production" = CStr(preactor.ReadFieldString("Orders", "Operation Name", orderRecordNum))) Then
                                '        i = 4
                                '        'ElseIf ("Production" = CStr(preactor.ReadFieldString("Orders", "Operation Name", parentOrderRecordNum))) Then
                                '        '    i = 4
                                '    Else
                                '        orderRecordNum = 0
                                '        i = i + 1
                                '    End If
                                'Next i

                                Dim oldSelectedBlockId As String = "0"
                                Dim oldSelectedBlockNo As String = "0"
                                oldSelectedBlockId = preactor.ReadFieldString("Orders", "K201 Selected Block Id", orderRecordNum)
                                If IsNumeric(oldSelectedBlockId) Then
                                    If CInt(oldSelectedBlockId) > 0 Then
                                        Dim orderCount As Integer = preactor.RecordCount("Orders")
                                        Dim x As Integer = 1
                                        Do
                                            Dim OrderNo_ = preactor.ReadFieldString("Orders", "Order No.", x)
                                            Dim string_Attribute_2 = preactor.ReadFieldString("Orders", "String Attribute 2", x)

                                            orderRcNum = preactor.FindMatchingRecord("Orders", "Order No.", orderRcNum, OrderNo_)
                                            oldSelectedBlockNo = preactor.ReadFieldString("Orders", "K201 Selected Block Id", orderRcNum)
                                            If IsNumeric(oldSelectedBlockNo) Then
                                                If CInt(oldSelectedBlockNo) > 0 Then
                                                    If CInt(oldSelectedBlockNo) = CInt(oldSelectedBlockId) Then
                                                        preactor.WriteField("Orders", "K201 Selected Block Id", orderRcNum, 0)
                                                        preactor.WriteField("Orders", "String Attribute 1", orderRecordNum, string_Attribute_2)


                                                    End If
                                                End If
                                            End If
                                            orderRcNum = 0
                                            x = x + 1
                                        Loop While x <= orderCount
                                    End If
                                End If
                                preactor.Commit("Orders")
                                preactor.Redraw()
                                ''While orderRecordNum > 0
                                If orderRecordNum > 0 Then
                                    If "Production" = CStr(preactor.ReadFieldString("Orders", "Operation Name", orderRecordNum)) Then
                                        Try
                                            Dim int_Resource As Integer
                                            int_Resource = preactor.ReadFieldInt("Orders", "Resource", orderRecordNum)
                                            'MsgBox(orderRecordNum)
                                            'MsgBox(int_Resource)

                                            Dim sequance As Integer = CInt(K201_GetOrderResourceDataSequance(connetionString, CInt(orderId), int_Resource))
                                            ''Dim sequance As Integer = CInt(K201_GetOrderResourceDataSequance(connetionString, orderRecordNum, int_Resource))
                                            orderNo = preactor.ReadFieldString("Orders", "Order No.", orderRecordNum)
                                            'preactor.WriteListField("Orders", "Resource Rate Per Hour", orderRecordNum, rate, sequance)

                                            ''Commented by  MilanAmarasooriya 20221011
                                            ''preactor.WriteListField("Orders", "Resource Rate Per Hour", orderRecordNum, (((formarRatio / totalAssignedFormerQty) - 0.001) * productResourceRate * totalAssignedFormerQty / line_Capacity), sequance)
                                            ''Commented by  MilanAmarasooriya 20221011
                                            preactor.WriteListField("Orders", "Resource Rate Per Hour", orderRecordNum, rate, sequance)


                                            preactor.WriteField("Orders", "K201 Selected Block Id", orderRecordNum, blockSelectionId)
                                            preactor.WriteField("Orders", "String Attribute 1", orderRecordNum, CStr(blockSelectionId))
                                            ''Change 20220406 by request sheriff
                                            ''preactor.WriteMatrixField("Orders", "Resource Constraint Qty", orderRecordNum, rate, sequance, 0)
                                            ''preactor.WriteMatrixField("Orders", "Resource Constraint Qty", orderRecordNum, ((formarRatio / totalAssignedFormerQty) - 0.001), sequance, 0)
                                            ''Commented by  MilanAmarasooriya 20221012
                                            ''preactor.WriteMatrixField("Orders", "Resource Constraint Qty", orderRecordNum, Math.Round(((formarRatio / totalAssignedFormerQty) - 0.006), 2), sequance, 0)
                                            ''Added by  MilanAmarasooriya 20221012
                                            preactor.WriteMatrixField("Orders", "Resource Constraint Qty", orderRecordNum, Math.Round(formarRatio / totalAssignedFormerQty, 2), sequance, 0)

                                            preactor.WriteField("Orders", "Numerical Attribute 5", orderRecordNum, formarRatio)

                                            ''preactor.WriteField("Orders", "K201_FormerRatio", orderRecordNum, ((formarRatio / totalAssignedFormerQty) - 0.001))
                                            ''Commented by  MilanAmarasooriya 20221012
                                            ''preactor.WriteField("Orders", "K201_FormerRatio", orderRecordNum, Math.Round(((formarRatio / totalAssignedFormerQty) - 0.006), 2))
                                            ''Added by  MilanAmarasooriya 20221012
                                            preactor.WriteField("Orders", "K201_FormerRatio", orderRecordNum, Math.Round(formarRatio / totalAssignedFormerQty, 2))

                                            ''Shriff's Request 20220407
                                            'Dim selectedConstraint As String = ""
                                            'selectedConstraint = preactor.ReadFieldString("Orders", "Selected Constraint 1", orderRecordNum)
                                            ''preactor.WriteField("Orders", "String Attribute 5", orderRecordNum, Secondary_Constraints)


                                            ''Name of the Secondary Constraints
                                            'Dim Secondary_Constraints As String = preactor.ReadFieldString("Orders", "Secondary Constraints", orderRecordNum, 1, 0)
                                            'preactor.WriteField("Orders", "String Attribute 5", orderRecordNum, Secondary_Constraints)
                                            'preactor.WriteMatrixField("Orders", "Constraint Quantity", orderRecordNum, formarRatio, 1, 0)

                                            ''Assign the entered former quantity into 'tot' variable
                                            'Dim pesp As IEventScriptsCore = EventScriptsFactory.CreateEventScriptCoreObject(preactorComObject, pespComObject)
                                            Dim matrixDim As MatrixDimensions
                                            matrixDim = preactor.MatrixFieldSize("Orders", "Secondary Constraints", orderRecordNum)

                                            For i = 1 To matrixDim.X
                                                'MsgBox(preactor.ReadFieldString("Orders", "Secondary Constraints", orderRecordNum, i))
                                                Dim Secondary_Constraints As String = preactor.ReadFieldString("Orders", "Secondary Constraints", orderRecordNum, i)
                                                If Secondary_Constraints.ToLower().Contains("tot") Then
                                                    preactor.WriteMatrixField("Orders", "Constraint Quantity", orderRecordNum, formarRatio, i, 0)
                                                Else
                                                    preactor.WriteMatrixField("Orders", "Constraint Quantity", orderRecordNum, 0, i, 0)
                                                    preactor.WriteField("Orders", "String Attribute 5", orderRecordNum, Secondary_Constraints)
                                                End If
                                            Next


                                            Dim order_No As String
                                            order_No = preactor.ReadFieldString("Orders", "Order No.", orderRecordNum)
                                            Dim op_No As Integer
                                            op_No = preactor.ReadFieldInt("Orders", "Op. No.", orderRecordNum)
                                            Dim resource As Integer
                                            resource = preactor.ReadFieldInt("Orders", "Resource", orderRecordNum)
                                            Dim startTime As DateTime
                                            startTime = preactor.ReadFieldDateTime("Orders", "Start Time", orderRecordNum)
                                            Dim productCode As String
                                            productCode = preactor.ReadFieldString("Orders", "String Attribute 4", orderRecordNum)

                                            Dim isSameProduct As Integer = CInt(K201_GetPreviousJobIsSameProductCodeforSelectedBlock_Sp(connetionString, order_No, op_No, resource, startTime, productCode))

                                            If isSameProduct = 1 Then
                                                preactor.WriteField("Orders", "Setup Time", orderRecordNum, 0)
                                            End If

                                        Catch ex As Exception
                                            MsgBox(ex.Message.ToString() + " - Value Writing",, "Error")
                                        End Try
                                    End If
                                End If
                                orderRecordNum = 0
                            Next
                            preactor.Commit("Orders")
                            preactor.Redraw()
                            planningboard.Close()
                            planningboard.SetLocateState(False)

                        End If
                    End If
                Else
                    preactor.Commit("Orders")
                    preactor.Redraw()
                    planningboard.Close()
                    planningboard.SetLocateState(False)
                End If

            Catch ex As Exception
                MsgBox(ex.Message.ToString() + " - FormarRatioCalculation",, "Error")
            End Try

        End If
        Return 1
    End Function

#End Region
    '' get K201_GetOrderResourceRate using execute K201_GetOrdersResourceRate_Sp
    Public Function K201_GetOrderResourceRate(ByRef connetionString As String, ByRef orerderNo As String, ByRef resourceId As Integer) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetOrdersResourceRate_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNum", orerderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@ResourceId", resourceId)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            Dim ResourceRatePerHour As Decimal = 0
            param = New SqlParameter("@ResourceRatePerHour", ResourceRatePerHour)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Orders resource rate not define",, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetProductResourceRate using execute K201_GetProductResourceRate_Sp
    Public Function K201_GetProductResourceRate(ByRef connetionString As String, ByRef orderNo As String) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetProductResourceRate_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter


            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim ProductResourceRate As Decimal = 0
            param = New SqlParameter("@ProductResourceRate", ProductResourceRate)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Product resource rate not define",, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetOrderResourceRate using execute K201_GetOrderResourceRate_Sp
    Public Function K201_GetOrderResourceRate(ByRef connetionString As String, ByRef orderNo As String) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetOrderResourceRate_Sp"
            command.CommandTimeout = 600
            Dim param As SqlParameter


            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim OrderResourceRate As Decimal = 0
            param = New SqlParameter("@OrderResourceRate", OrderResourceRate)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Order resource rate not found" + ex.Message,, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetFormersPerPlatoon using execute K201_GetFormersPerPlatoon_Sp
    Public Function K201_GetFormersPerPlatoon(ByRef connetionString As String, ByRef orderNo As String) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetFormersPerPlatoon_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter


            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim availableFormer As Decimal = 0
            param = New SqlParameter("@FormersPerPlatoon", availableFormer)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Formers per platoon not define",, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetExtraFormerTbl using execute K201_GetExtraFormers_Sp
    Public Function K201_GetExtraFormerTbl(ByRef connetionString As String, ByRef orderNos As String) As DataTable

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetExtraFormers_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNos", orderNos)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)

            Return ds.Tables(0)

            connection.Close()
        Catch ex As Exception
            MsgBox("Extra formers not define",, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetAvilableFormer using execute K201_GetAvailableFormer_Sp
    Public Function K201_GetAvilableFormer(ByRef connetionString As String, ByRef orderNo As String) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetAvailableFormer_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim availableFormer As Decimal = 0
            param = New SqlParameter("@AvailableFormer", availableFormer)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Available former not define",, "Error")
        Finally

        End Try
    End Function

    '' get K201_GetTotalAvilableFormer using execute K201_GetTotalAvailableFormer_Sp
    Public Function K201_GetTotalAvilableFormer(ByRef connetionString As String, ByRef orderNo As String) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetTotalAvailableFormer_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            Dim availableFormer As Decimal = 0
            param = New SqlParameter("@AvailableFormer", availableFormer)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Decimal
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CDec(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Secondary Constarints Starting with 'tot' not defined",, "Error")
        Finally

        End Try
    End Function


    '' get K201_GenarateNewBlockSelectionIdFormer using execute K201_GenarateNewBlockSelectionId_Sp
    Public Function K201_GenarateNewBlockSelectionIdFormer(ByRef connetionString As String) As Integer

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GenarateNewBlockSelectionId_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@BlockId", 1)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (param.Value.ToString = "0") Then
                Return CInt(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Block Id Genarate fail",, "Error")
        Finally

        End Try

    End Function
    '' get K201_GetDueDateExceedOrders 
    Public Function K201_GetDueDateExceedOrders(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Try
            Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
            Dim planningboard As IPlanningBoard = preactor.PlanningBoard
            Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
            Dim dueDateExcJobTbl As DataTable = New DataTable()
            Dim oFormDueDateExcJob As New DueDateExceededJobDetails()

            dueDateExcJobTbl = K201_GetOrderResourceRate(connetionString)
            oFormDueDateExcJob.tblDueDateExcJob = dueDateExcJobTbl
            oFormDueDateExcJob.connetionString = connetionString
            oFormDueDateExcJob.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function
    '' get K201_GetOrderResourceRate using execute K201_GetDueDateExceededJobs_Sp and return resourcerate table
    Public Function K201_GetOrderResourceRate(ByRef connetionString As String) As DataTable

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetDueDateExceededJobs_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            Dim tblDueDateExcJob As New DataTable("MyTable")
            adapter.Fill(tblDueDateExcJob)
            connection.Close()
            Return tblDueDateExcJob
        Catch ex As Exception
            MsgBox("Orders resource rate not define",, "Error")
        Finally
        End Try

    End Function
    '' get K201_GetOrderResourceDataSequance using execute K201_GetOrderResourceDataSequance_Sp
    Public Function K201_GetOrderResourceDataSequance(ByRef connetionString As String, ByRef orderId As Integer, ByRef resourceId As Integer) As Integer

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)
            'MsgBox(connection)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetOrderResourceDataSequance_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderId", orderId)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            param = New SqlParameter("@ResourceId", resourceId)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            Dim sequance As Decimal = 0
            param = New SqlParameter("@Sequance", sequance)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()


            If Not (param.Value.ToString = "") Then
                If Not (param.Value.ToString = "0") Then
                    Return CInt(param.Value)
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message + "Resource Data Sequence Error",, "Error")
        Finally

        End Try

    End Function

    '' get K201_GetPreviousJobIsSameProductCode using execute K201_GetPreviousJobIsSameProductCode_Sp
    Public Function K201_GetPreviousJobIsSameProductCodeforSelectedBlock_Sp(ByRef connetionString As String, ByRef orderNo As String, ByRef opNo As Integer, resourceId As Integer, startTime As DateTime, productCode As String) As Integer

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetPreviousJobIsSameProductCodeforSelectedBlock_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@OrderNo", orderNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@OpNo", opNo)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            param = New SqlParameter("@ResourceId", resourceId)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            param = New SqlParameter("@StartTime", startTime)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.DateTime
            command.Parameters.Add(param)

            param = New SqlParameter("@ProductCode", productCode)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@IsSameProduct", 1)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            If Not (CInt(param.Value) = 0) Then
                Return CInt(param.Value)
            Else
                Return 0
            End If
            connection.Close()
        Catch ex As Exception
            MsgBox("Previous product code error",, "Error")
        Finally

        End Try

    End Function

    ''K201_BatchSpilt programe will execuate when user write click  and execute K201_BatchSpilt 
    Public Function K201_BatchSpilt(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard

        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim dt As DataTable = New DataTable()

        Dim c_OrderNo As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        Dim c_ProductCode As DataColumn = New DataColumn("ProductCode", Type.[GetType]("System.String"))
        Dim c_OrderQuantity As DataColumn = New DataColumn("ProductQuantity", Type.[GetType]("System.Double"))
        dt.Columns.Add(c_OrderNo)
        dt.Columns.Add(c_OrderQuantity)
        dt.Columns.Add(c_OrderQuantity)

        Dim num As Integer = preactor.RecordCount("Orders")

        Dim x As Integer = 1
        Dim firstRecord As Integer = 1

        Do
            If (planningboard.GetOperationLocateState(x)) Then
                Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", x)
                If strBelongsToOrderNo = "PARENT" Then
                    Dim dr As DataRow = dt.NewRow()
                    dr("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", x)
                    dr("ProductCode") = preactor.ReadFieldString("Orders", "Part No.", x)
                    dr("ProductQuantity") = preactor.ReadFieldString("Orders", "Quantity", x)
                    dt.Rows.Add(dr)
                End If
            End If
            x = x + 1
        Loop While x <= num
    End Function
#End Region

#Region "K201_AddCalendarException"
    '' get K201_AddCalendarException
    Public Function K201_AddCalendarException(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim exceptionResources As DataTable
        exceptionResources = K201_GetExceptionResources(connetionString)

        Dim resRowId As Integer

        For Each exres As DataRow In exceptionResources.Rows
            resRowId = preactor.FindMatchingRecord("Resources", "Name", resRowId, CStr(exres("Resource")))
            If resRowId > 0 Then
                planningboard.CreatePrimaryCalendarException(resRowId, CDate(exres("StartTime")), CDate(exres("EndTime")), "No Order for Size")
                preactor.Commit("Orders")
            End If
        Next

        preactor.Commit("Orders")
        preactor.Redraw()
        planningboard.Close()

        Return 1
    End Function
    '' get K201_GetExceptionResources using execute K201_GetExceptionResources_Sp
    Public Function K201_GetExceptionResources(ByRef connetionString As String) As DataTable

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetExceptionResources_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            Dim tblDueDateExcJob As New DataTable("MyTable")
            adapter.Fill(tblDueDateExcJob)

            connection.Close()
            Return tblDueDateExcJob
        Catch ex As Exception
            MsgBox("Orders resource rate not define",, "error")
        Finally

        End Try

    End Function
    '' genarate Bulk Execptions
    Public Function K201_BulkExecptionGenarate(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim exceptionResources As DataTable
        exceptionResources = K201_GetBulkExceptionResources(connetionString)

        Dim resRowId As Integer
        Dim count As Integer = 1
        Dim strBooks(4, 1) As String

        For Each exres As DataRow In exceptionResources.Rows
            resRowId = preactor.FindMatchingRecord("Resources", "Name", resRowId, CStr(exres("Resource")))
            If resRowId > 0 Then
                planningboard.CreatePrimaryCalendarException(resRowId, CDate(exres("StartTime")), CDate(exres("EndTime")), "No Order for Size")
                count = count + 1
            End If
            resRowId = 0
        Next
        preactor.Commit("Orders")
        preactor.Redraw()
        planningboard.Close()

        Return 1
    End Function
    '' get K201_GetBulkExceptionResources using execute K201_GetBulkExceptionResources_Sp
    Public Function K201_GetBulkExceptionResources(ByRef connetionString As String) As DataTable
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_GetBulkExceptionResources_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()
            Dim tblDueDateExcJob As New DataTable("MyTable")
            adapter.Fill(tblDueDateExcJob)

            connection.Close()
            Return tblDueDateExcJob
        Catch ex As Exception
            MsgBox("Orders resource rate not define",, "error")
        Finally

        End Try
    End Function


#End Region

#Region "K201_DeleteCalendarException"
    ''K201_DeleteCalendarException programme
    Public Function K201_DeleteCalendarException(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")

        Dim orderdistinct_dt As DataTable = New DataTable()
        Dim order_ss As DataColumn = New DataColumn("OrderNo", Type.[GetType]("System.String"))
        orderdistinct_dt.Columns.Add(order_ss)
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim i As Integer = 1
        Do
            If (planningboard.GetOperationLocateState(i)) Then

                Dim strBelongsToOrderNo As String = preactor.ReadFieldString("Orders", "Belongs to Order No.", i)

                If strBelongsToOrderNo = "PARENT" Then
                    Dim dt_sr As DataRow = orderdistinct_dt.NewRow()
                    dt_sr("OrderNo") = preactor.ReadFieldString("Orders", "Order No.", i)
                    orderdistinct_dt.Rows.Add(dt_sr)
                End If
            End If

            i = i + 1
        Loop While i <= num
        orderdistinct_dt = orderdistinct_dt.DefaultView.ToTable()
        ''create table varible and asign column and values
        Dim orderRes_dt As DataTable = New DataTable()
        Dim orderno_ss As DataColumn = New DataColumn("Order", Type.[GetType]("System.String"))
        Dim resource_ss As DataColumn = New DataColumn("Resource", Type.[GetType]("System.String"))
        Dim resourceStartTime_ss As DataColumn = New DataColumn("ResourceStartTime", Type.[GetType]("System.String"))
        Dim resourceEndTime_ss As DataColumn = New DataColumn("ResourceEndTime", Type.[GetType]("System.String"))
        Dim resourceId_ss As DataColumn = New DataColumn("ResourceRowId", Type.[GetType]("System.String"))

        orderRes_dt.Columns.Add(orderno_ss)
        orderRes_dt.Columns.Add(resource_ss)
        orderRes_dt.Columns.Add(resourceStartTime_ss)
        orderRes_dt.Columns.Add(resourceEndTime_ss)
        orderRes_dt.Columns.Add(resourceId_ss)

        Dim orderNo As String
        Dim resRowId As Integer
        Dim ResourceStartTime As DateTime
        Dim ResourceEndTime As DateTime
        Dim ResourceRowId As Integer
        For Each orderr As DataRow In orderdistinct_dt.Rows
            orderNo = orderr("OrderNo").ToString()
            Dim order_Num As Integer = 0
            order_Num = preactor.FindMatchingRecord("Orders", "Order No.", order_Num, orderNo)
            Dim resourceName As String = preactor.ReadFieldString("Orders", "Resource", order_Num)

            resRowId = preactor.FindMatchingRecord("Resources", "Name", resRowId, resourceName)

            If resRowId > 0 Then

                Dim actualResource As String = preactor.ReadFieldString("Resources", "K201 Actual Resource", resRowId)

                Dim actualResourceRecordNo As Integer = 0
                Try

                    actualResourceRecordNo = preactor.FindMatchingRecord("Resources", "K201 Actual Resource", actualResourceRecordNo, actualResource)
                    While actualResourceRecordNo > 0
                        Dim ors As DataRow = orderRes_dt.NewRow()
                        Dim rName As String = preactor.ReadFieldString("Resources", "Name", actualResourceRecordNo)

                        ors("Order") = order_Num
                        ors("Resource") = rName
                        ors("ResourceRowId") = actualResourceRecordNo
                        ors("ResourceStartTime") = preactor.ReadFieldString("Orders", "Start Time", order_Num)
                        ors("ResourceEndTime") = preactor.ReadFieldString("Orders", "End Time", order_Num)

                        orderRes_dt.Rows.Add(ors)
                        actualResourceRecordNo = preactor.FindMatchingRecord("Resources", "K201 Actual Resource", actualResourceRecordNo, actualResource)

                    End While
                Catch ex As Exception
                    MsgBox("Actual Resource Not Define",, "Error")
                End Try

            End If
        Next
        resRowId = 0
        For Each orderr As DataRow In orderRes_dt.Rows

            ResourceStartTime = CDate(orderr("ResourceStartTime"))
            ResourceEndTime = CDate(orderr("ResourceEndTime"))
            ResourceRowId = CInt(orderr("ResourceRowId"))

            K201_DeleteCalendarExceptionByResourceAndDaterange(connetionString, CStr(orderr("Resource")), ResourceStartTime, ResourceEndTime)
            preactor.Commit("Orders")
            planningboard.UpdatePrimaryResourceCalendar(ResourceRowId)

        Next
        preactor.Redraw()
        planningboard.Close()
        planningboard.SetLocateState(False)

        Return 1
    End Function
    '' get K201_DeleteCalendarExceptionByResourceAndDaterange using execute K201_DeleteCalendarException_Sp
    Public Function K201_DeleteCalendarExceptionByResourceAndDaterange(ByRef connetionString As String, ByRef resource As String, startDate As Date, endDate As Date) As Decimal

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_DeleteCalendarException_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@Resource", resource)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.String
            command.Parameters.Add(param)

            param = New SqlParameter("@StartDate", startDate)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Date
            command.Parameters.Add(param)

            Dim availableFormer As Decimal = 0
            param = New SqlParameter("@EndDate", endDate)
            param.Direction = ParameterDirection.Input
            param.DbType = DbType.Date
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("Delete calendar exception error",, "error")
        Finally

        End Try

    End Function
#End Region

#Region "Show Late Order"
    ''Show Late Order
    Public Function K201_ShowLateOrders(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim d As Integer = 1
        Dim orderNum As String
        Dim orderEndDate As String
        Dim dueDate As String
        Dim strMessage As String
        Do
            orderNum = preactor.ReadFieldString("Orders", "Order No.", d)
            orderEndDate = preactor.ReadFieldString("Orders", "End Time", d)
            dueDate = preactor.ReadFieldString("Orders", "Due Date", d)
            If Not ((String.IsNullOrEmpty(dueDate)) Or dueDate = "Unspecified" Or (String.IsNullOrEmpty(orderEndDate)) Or orderEndDate = "Unspecified") Then
                If CDate(orderEndDate) > CDate(dueDate) Then
                    strMessage = strMessage + (IIf(strMessage = "", orderNum, " , " + orderNum)).ToString
                End If
            End If
            d = d + 1
        Loop While d <= num
        If Not (strMessage = "") Then
            MsgBox("Late Orders " + strMessage, vbInformation, "APS")
        End If
    End Function
#End Region
#Region "Show Late Order"
    ''K201_BOMShortageCalculation
    Public Function K201_BOMShortageCalculation(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)

            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_BOMShortagesCalculation_Sp"
            command.CommandTimeout = 600

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()

            connection.Close()
        Catch ex As Exception
            MsgBox("BOM Shortages Calculation Faild",, "Error")
        Finally

        End Try
        MsgBox("BOM Shortages Calculated",, "Information")
        Return 0
    End Function
#End Region
#Region "Resource Group Change"
    Public Function K201_ResourceGroupChange(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim numo As Integer = preactor.RecordCount("Orders")
        MsgBox("ResourceGroup Change")
        Return 0
    End Function
#End Region


    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef K201_AssignValue As Integer) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim numo As Integer = preactor.RecordCount("Orders")
        Dim num As Integer = preactor.RecordCount("Demand")
        'MsgBox(numo)
        K201_AssignValue = 5

        Return K201_AssignValue
    End Function
    Public Function CSVExport(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim stat As Int16 = 0

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)
            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_CSVExport_Sp"
            Dim param As SqlParameter
            command.CommandTimeout = 340
            adapter = New SqlDataAdapter(command)
            Dim ds As DataSet = New DataSet()
            adapter.Fill(ds)
            connection.Close()

            Dim strData As String
            Dim fileloc As String = "D:\Milan\ImportFile\PrimaryCalendarPeriods.csv"
            Dim datatbl As DataTable = ds.Tables(0)
            Dim isfirstrow As Integer = 1
            For Each row As DataRow In datatbl.Rows
                Dim line As String = ""
                If isfirstrow = 1 Then
                    For Each column As DataColumn In datatbl.Columns
                        line += "," & (column.ColumnName).ToString()
                    Next
                    strData += line.Substring(1) & vbCrLf
                    isfirstrow = 0
                    line = ""
                End If

                For Each column As DataColumn In datatbl.Columns
                    line += "," & row(column.ColumnName).ToString()
                Next
                strData += line.Substring(1) & vbCrLf
            Next

            If File.Exists(fileloc) Then
                File.Delete(fileloc)
            End If
            Using sw As StreamWriter = New StreamWriter(fileloc)
                sw.WriteLine(strData)
            End Using
        Catch ex As Exception
            stat = -1
            MsgBox("Demand date export error plase contact administrator...",, "error")
        Finally
            IIf(stat = 1, stat = 0, MsgBox("Demand Export Complted",, "Information"))
        End Try

        Return 0
    End Function



    Public Function K201_MySqlDataValidation_Sp(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        ''define variable and assign
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")

        Try
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As New SqlCommand

            connection = New SqlConnection(connetionString)



            connection.Open()
            command.Connection = connection
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = "K201_Int_MySqlDataValidation_Sp"
            command.CommandTimeout = 600

            Dim param As SqlParameter

            param = New SqlParameter("@Result", 1)
            param.Direction = ParameterDirection.Output
            param.DbType = DbType.Int16
            command.Parameters.Add(param)

            adapter = New SqlDataAdapter(command)
            command.ExecuteNonQuery()


            If (CInt(param.Value) = 1) Then
                MsgBox("Some Data Needs to be Validated, Contact IT Suppport.",, "Data Validation")
            End If

            Return 0
            connection.Close()
        Catch ex As Exception
            MsgBox("Error while execute K201_MySqlDataValidation_Sp" + "|" + ex.Message,, "Error")
        Finally

        End Try

    End Function


    Public Function K201_UpdateProductionLineOperation(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        Dim connetionString As String = preactor.ParseShellString("{DB CONNECT STRING}")
        Dim y As Integer = 1

        Dim num As Integer = preactor.RecordCount("Orders")
        Dim K201_ProductionLineOperationRecordId As Integer
        Dim subOrder As Integer = 0
        Do
            If (planningboard.GetOperationLocateState(y)) Then
                Dim intOperationNumber As Integer = preactor.ReadFieldInt("Orders", "Op. No.", y)
                If intOperationNumber = 10 Then

                    Dim orno As String = preactor.ReadFieldString("Orders", "Order No.", y)
                    Dim opno As Integer = preactor.ReadFieldInt("Orders", "Op. No.", y)




                    Dim resourceName As String = preactor.ReadFieldString("Orders", "Resource", y)
                    If Not ((resourceName = "Nothing") Or (resourceName = "Unspecified")) Then
                        K201_ProductionLineOperationRecordId = 0
                        K201_ProductionLineOperationRecordId = preactor.FindMatchingRecord("K201_ProductionLineOperation", "Line", K201_ProductionLineOperationRecordId, resourceName)

                        If K201_ProductionLineOperationRecordId > 0 Then
                            subOrder = 0
                            subOrder = preactor.FindMatchingRecord("Orders", "Order No.", subOrder, orno)

                            While subOrder > 0
                                If subOrder > 0 Then
                                    Dim op_no As Integer = preactor.ReadFieldInt("Orders", "Op. No.", subOrder)
                                    'If op_no = 10 Then
                                    '    Dim Dipping As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Dipping", K201_ProductionLineOperationRecordId)
                                    '    If Not Dipping = 1 Then
                                    '        preactor.WriteField("Orders", "Disable Operation", y, 0)
                                    '    Else
                                    '        preactor.WriteField("Orders", "Disable Operation", y, 0)
                                    '    End If
                                    'End If
                                    If op_no = 30 Then
                                        Dim Chlorination As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Chlorination", K201_ProductionLineOperationRecordId)
                                        If Chlorination = 1 Then
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                        Else
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                        End If
                                    End If
                                    If op_no = 50 Then
                                        Dim Printing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Printing", K201_ProductionLineOperationRecordId)
                                        If Printing = 1 Then
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                        Else
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                        End If
                                    End If
                                    If op_no = 60 Then
                                        Dim Packing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Packing", K201_ProductionLineOperationRecordId)
                                        If Packing = 1 Then
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                        Else
                                            preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                        End If
                                    End If
                                    preactor.Commit("Orders")
                                End If
                                subOrder = preactor.FindMatchingRecord("Orders", "Order No.", subOrder, orno)
                            End While

                        End If
                        K201_ProductionLineOperationRecordId = preactor.FindMatchingRecord("K201_ProductionLineOperation", "Line", K201_ProductionLineOperationRecordId, resourceName)
                    End If

                End If
            End If
            y = y + 1
        Loop While y <= num
        ''preactor.Commit("Orders")

    End Function
    Public Function K201_GantChartBlockDropAndUpdateOrder(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef i As Integer) As Integer
        ''Check Resource and Set hold the order
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Dim resourceName As String = preactor.ReadFieldString("Orders", "Resource", i)
        Dim op_noMain As Integer = preactor.ReadFieldInt("Orders", "Op. No.", i)
        Dim orno As String = preactor.ReadFieldString("Orders", "Order No.", i)

        Dim subOrder As Integer = 0
        subOrder = preactor.FindMatchingRecord("Orders", "Order No.", subOrder, orno)

        If op_noMain = 10 Then
            If Not ((resourceName = "Nothing") Or (resourceName = "Unspecified")) Then
                Dim K201_ProductionLineOperationRecordId As Integer = preactor.FindMatchingRecord("K201_ProductionLineOperation", "Line", K201_ProductionLineOperationRecordId, resourceName)
                If K201_ProductionLineOperationRecordId > 0 Then
                    While subOrder > 0
                        If subOrder > 0 Then
                            Dim op_no As Integer = preactor.ReadFieldInt("Orders", "Op. No.", subOrder)
                            'If op_no = 10 Then
                            '    Dim Dipping As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Dipping", K201_ProductionLineOperationRecordId)
                            '    If Not Dipping = 1 Then
                            '        preactor.WriteField("Orders", "Disable Operation", i, 0)
                            '    Else
                            '        preactor.WriteField("Orders", "Disable Operation", i, 0)
                            '    End If
                            'End If
                            If op_no = 30 Then
                                    Dim Chlorination As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Chlorination", K201_ProductionLineOperationRecordId)
                                If Chlorination = 1 Then
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                Else
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                End If
                                End If
                                If op_no = 50 Then
                                Dim Printing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Printing", K201_ProductionLineOperationRecordId)
                                If Printing = 1 Then
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                Else
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                End If
                            End If
                            If op_no = 60 Then
                                Dim Packing As Integer = preactor.ReadFieldInt("K201_ProductionLineOperation", "Packing", K201_ProductionLineOperationRecordId)
                                If Packing = 1 Then
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 1)
                                Else
                                    preactor.WriteField("Orders", "Disable Operation", subOrder, 0)
                                End If
                            End If
                            preactor.Commit("Orders")

                        End If
                        subOrder = preactor.FindMatchingRecord("Orders", "Order No.", subOrder, orno)
                    End While
                End If
            End If
            preactor.Commit("Orders")
        End If

        ''Check Resource and Set hold the order
    End Function


    Public Function K201_ChangeDamagePercentage(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object, ByRef RecordNumber As Integer) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim damagePercentageTxt As Double
        Dim num As Integer = preactor.RecordCount("Orders")
        Dim changeDamagePercentageForm As New DamagePercentageForm()
        changeDamagePercentageForm.ShowDialog()

        If changeDamagePercentageForm.isOkClick = True Then


            If Not changeDamagePercentageForm.TextBox1.Text = "" Then
                damagePercentageTxt = CDec(changeDamagePercentageForm.TextBox1.Text)
                Dim quantity As Integer = preactor.ReadFieldInt("Orders", "K201 Original Order Quantity", RecordNumber)
                Dim damageQty As Double = quantity * damagePercentageTxt / 100
                Dim newQty As Double = Math.Ceiling(quantity + damageQty)

                Dim orderNo As String = preactor.ReadFieldString("Orders", "Order No.", RecordNumber)
                Dim i As Integer = 1
                Do
                    Dim newStrOrderNo As String = preactor.ReadFieldString("Orders", "Order No.", i)
                    If orderNo = newStrOrderNo Then
                        preactor.WriteField("Orders", "Numerical Attribute 3", i, damagePercentageTxt)
                        preactor.WriteField("Orders", "Quantity", i, newQty)
                    End If
                    i = i + 1
                Loop While i <= num

            Else
                MsgBox("Please enter damage percentage")
            End If


        End If
        preactor.Commit("Orders")
        Return 0
    End Function

    Public Function K201_OrderDeleteOk(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        Dim planningboard As IPlanningBoard = preactor.PlanningBoard
        preactor.Commit("Orders")
        preactor.Load("Orders", "Schedule")
        preactor.Commit("Orders")
        ''preactor.SortDataTable("Orders", "Order No.")
        MsgBox("Orders deletion unauthorized....")

        Return 0
    End Function
    Public Function K201_GetCurrentUser(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer
        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)
        preactor.Clear("K201_DataImportUserDetails")
        preactor.Commit("K201_DataImportUserDetails")

        ''Dim userName As String = preactor.ParseShellString("{USER NAME}")
        Dim userName As String = Environment.UserName
        Dim pcDate As DateTime = System.DateTime.Now()
        MsgBox(userName + " " + CStr(pcDate))
        Dim newBlock As Integer = preactor.CreateRecord("K201_DataImportUserDetails")
        Dim newRecordNum As Integer = preactor.ReadFieldInt("K201_DataImportUserDetails", "Number", newBlock)
        preactor.WriteField("K201_DataImportUserDetails", "Number", newBlock, newRecordNum)
        preactor.WriteField("K201_DataImportUserDetails", "ExecutedUser", newBlock, userName)
        preactor.WriteField("K201_DataImportUserDetails", "ExecutedTime", newBlock, pcDate)
        preactor.Commit("K201_DataImportUserDetails")
        '' MsgBox(userName)
        Return 0
    End Function

End Class
