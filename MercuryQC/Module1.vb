Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Module Module1
    Dim reportDir As String = String.Empty
    Sub Main()
        Dim bDirectoryCheck As Boolean = True, clientID As Integer = 0, strReportType As String = "Element"
        Dim clArgs() As String = Environment.GetCommandLineArgs()


        If Not Directory.Exists(My.Settings.LogDir) Then
            System.Environment.Exit(1)
        End If
        If clArgs.Length = 1 Then
            LogError("client ID not valid")
            System.Environment.Exit(1)
        End If
        clientID = clArgs(1)
        If Not Directory.Exists(My.Settings.TemplateDir) Then
            LogError("Template Directory does not exist")
            bDirectoryCheck = False
        End If
        If Not Directory.Exists(My.Settings.ReportDir) Then
            LogError("Report Directory does not exist")
            bDirectoryCheck = False
        End If
        If bDirectoryCheck = False Then
            System.Environment.Exit(1)
        End If
        If clArgs.Length = 3 Then
            strReportType = clArgs(2)
        End If
        reportDir = My.Settings.ReportDir & GetClientName(clientID, My.Settings.ConnectionString) & "\QC\"
        Select Case strReportType
            Case "Element"
                runElementReport(clientID)
            Case "Mailed"
                runMailedReport(clientID)
            Case "CampaignMatch"
                runCampaignMatchReport(clientID)
            Case "OrderSummary"
                runOrderSummaryReport(clientID)
            Case "MatchLevel"
                runMatchLevelReport(clientID)
            Case "MailedData"
                runMailedDataReport(clientID)
            Case Else

        End Select

    End Sub

    Function ExecuteCMD(ByRef CMD As SqlCommand, connectionString As String) As DataSet

        Dim ds As New DataSet()

        Try
            Using connection As New SqlConnection(connectionString)
                CMD.Connection = connection

                'Assume that it's a stored procedure command type if there is no space in the command text. Example: "sp_Select_Customer" vs. "select * from Customers"
                If CMD.CommandText.Contains(" ") Then
                    CMD.CommandType = CommandType.Text
                Else
                    CMD.CommandType = CommandType.StoredProcedure
                End If

                Dim adapter As New SqlDataAdapter(CMD)
                adapter.SelectCommand.CommandTimeout = 300

                'fill the dataset
                adapter.Fill(ds)
                connection.Close()
            End Using
        Catch ex As Exception
            ' The connection failed. Display an error message.
            Throw New Exception("Database Error:  " & ex.Message)
        End Try

        Return ds
    End Function
    Private Function GetCurrentDataset(clientID As String, connectionString As String) As String
        Dim strRet As String = String.Empty
        Using conn As SqlConnection = New SqlConnection(connectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand("Select Top 1 MercuryDatabaseName from mrtDataMart1.mrtMeritSharedDatamart.[MERIT_MATCH].[MercuryProjects] where ClientID = " & clientID.ToString & "
            and MercuryProjectStatus = 'C'
            Order by MercuryProjectID Desc")
                cmd.Connection = conn
                strRet = cmd.ExecuteScalar()
            End Using
        End Using
        Return strRet
    End Function

    Private Sub LogError(sText As String)
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append(DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " " + sText)
        File.AppendAllText(My.Settings.LogDir + My.Application.Info.AssemblyName + ".txt", sb.ToString() + vbCrLf)
        sb.Clear()
    End Sub
    Private Sub runElementReport(clientID As Integer)
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "ElementQC_Template.xlsx")
        Try
            Dim strDataset As String = GetCurrentDataset(clientID, My.Settings.ConnectionString)
            Dim i As Integer = 2
            Using pck As ExcelPackage = New ExcelPackage(fi)
                Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets("Sheet1")
                Dim cmd As New SqlClient.SqlCommand
                cmd.CommandText = "Exec MercuryAdmin.dbo.p_ElementQCReport " & clientID.ToString
                Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)

                    For Each dr As DataRow In dt.Rows
                        For nCol As Integer = 0 To dt.Columns.Count - 1

                            Select Case dt.Columns(nCol).ColumnName
                                Case "CountPVS", "CountCUR", "ChangeCount", "ColumnIndex"
                                    wsQC.Cells(i, nCol + 1).Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* "" 0 ""??_);_(@_)"
                                    wsQC.SetValue(i, nCol + 1, CInt(dr.Item(nCol)))
                                    'Case "ChangePct", "ColumnPctPVS", "ColumnPctCUR"
                                    '    wsQC.Cells(i, nCol + 1).Style.Numberformat.Format = "0.00%"
                                    '    wsQC.SetValue(i, nCol + 1, CDbl(dr.Item(nCol)))
                                Case Else
                                    wsQC.SetValue(i, nCol + 1, dr.Item(nCol).ToString)
                            End Select




                        Next
                        i += 1
                    Next
                End Using
                wsQC.Name = strDataset
                Dim fiOut As FileInfo = New FileInfo(reportDir & "ElementQCReport_" & strDataset.Replace("Mercury_V4_", "") & ".xlsx")
                pck.SaveAs(fiOut)
            End Using
            System.Environment.Exit(0)
        Catch ex As Exception
            LogError(ex.ToString)
        End Try
    End Sub
    Private Sub runMailedReport(clientID As Integer)
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "MailedQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets("PivotData")
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_MailedQCReport " & clientID.ToString
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case dt.Columns(nCol).ColumnName
                            Case "Mailed"
                                wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                            Case Else
                                wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                        End Select
                    Next
                    nRow += 1
                Next


            End Using
            nRow -= 1
            '  Dim Range = wsQC.Workbook.Names("CampaignData")
            '  Range.Address = "A1:F" & nRow.ToString
            '  wsQC.Workbook.Names.Remove("CampaignData")
            ' wsQC.Workbook.Names.Add("CampaignData", Range)

            Dim fiOut As FileInfo = New FileInfo(reportDir & "MailedQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            pck.SaveAs(fiOut)
        End Using


    End Sub
    Private Sub runCampaignMatchReport(clientID As Integer)
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "CampaignMatchQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets("PivotData")
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_CampaignMatchQCReport " & clientID.ToString
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case dt.Columns(nCol).ColumnName
                            Case "Mailed", "Response"
                                wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                            Case "ResponseDollars"
                                wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                            Case Else
                                wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                        End Select
                    Next
                    nRow += 1
                Next


            End Using
            nRow -= 1
            '  Dim Range = wsQC.Workbook.Names("CampaignData")
            '  Range.Address = "A1:F" & nRow.ToString
            '  wsQC.Workbook.Names.Remove("CampaignData")
            ' wsQC.Workbook.Names.Add("CampaignData", Range)

            Dim fiOut As FileInfo = New FileInfo(reportDir & "CampaignMatchQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            pck.SaveAs(fiOut)
        End Using


    End Sub
    Private Sub runOrderSummaryReport(clientID As Integer)
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "OrderSummaryQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 3
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets(1)
            wsQC.DeleteRow(3, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_OrderSummaryQCReport " & clientID.ToString
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case dt.Columns(nCol).ColumnName
                            Case "OrderCount", "PrevOrderCount"
                                If Not IsDBNull(dr.Item(nCol)) Then
                                    wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "#,##0"
                                End If
                                wsQC.Cells(nRow, nCol + 1).Style.Border.Left.Style = ExcelBorderStyle.Medium
                            Case "Revenue", "PrevRevenue", "AOV", "PrevAOV"
                                If Not IsDBNull(dr.Item(nCol)) Then
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "$#,##0"
                                End If
                            Case "Variance"
                                If Not IsDBNull(dr.Item(nCol)) Then
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "0%"
                                End If
                                wsQC.Cells(nRow, nCol + 1).Style.Border.Left.Style = ExcelBorderStyle.Medium
                            Case Else
                                wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                        End Select
                    Next
                    nRow += 1
                Next


            End Using


            wsQC.Cells(wsQC.Dimension.Address).AutoFitColumns()
            Dim fiOut As FileInfo = New FileInfo(reportDir & "OrderSummaryQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            pck.SaveAs(fiOut)
        End Using


    End Sub
    Private Sub runMatchLevelReport(clientID As Integer)
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "MatchLevelQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 3
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets(1)
            wsQC.DeleteRow(3, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_MatchLevelQCReport " & clientID.ToString
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case dt.Columns(nCol).ColumnName
                            Case "CurGrossMatchCount", "CurNetMatchCount", "PrevGrossMatchCount", "PrevNetMatchCount"
                                wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                                wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "#,##0"
                                wsQC.Cells(nRow, nCol + 1).Style.Border.Left.Style = ExcelBorderStyle.Medium
                                wsQC.Cells(nRow, nCol + 1).Style.Border.Right.Style = ExcelBorderStyle.Medium
                            Case "Variance"
                                wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "0%"
                            Case Else
                                wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                        End Select

                    Next
                    nRow += 1
                Next


            End Using
            Dim fiOut As FileInfo = New FileInfo(reportDir & "MatchLevelQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            pck.SaveAs(fiOut)
        End Using


    End Sub
    Private Sub runMailedDataReport(clientID As Integer)
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "MailedDataQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets(1)
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_MailedDataQCReport " & clientID.ToString
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    If Not dr.Item(0).ToString.Contains("Sub-Total") Then 'skip SubTotals
                        For nCol As Integer = 0 To dt.Columns.Count - 2 'skip total sort
                            Select Case dt.Columns(nCol).ColumnName
                                Case "Mailed", "Response", "Auxiliary Orders", "$/M Index", "RR Index", "AOV Index", "Aux$/Resp Index", "Aux Order/Resp Index"
                                    wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "#,##0"
                                Case "Response Rate"
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "0.00%"
                                Case "Response Dollars", "$/M", "AOV", "Auxiliary Dollars", "Aux$/Response", "GM", "AdCost", "OrderProcessing", "Contribution"
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "\$#,##0"
                                Case "$/Piece", "Aux Orders/Response", "Con/Order"
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "\$#,###.00"
                                Case "Aux Orders/Response"
                                    wsQC.SetValue(nRow, nCol + 1, CDbl(dr.Item(nCol)))
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "#.00"
                                Case Else
                                    wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                            End Select

                        Next
                        nRow += 1
                    End If
                Next


            End Using
            wsQC.Cells(wsQC.Dimension.Address).AutoFitColumns()
            Dim fiOut As FileInfo = New FileInfo(reportDir & "MailedDataQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            pck.SaveAs(fiOut)
        End Using


    End Sub
    Private Function GetClientName(clientID As String, connectionString As String) As String
        Dim strRet As String = String.Empty
        Using conn As SqlConnection = New SqlConnection(connectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand("Select Top 1 Cipher from mrtDatamart1.[mrtMeritSharedDatamart].[MERIT_MATCH].[Clients] where ClientID = " & clientID.ToString)
                cmd.Connection = conn
                strRet = cmd.ExecuteScalar()
            End Using
        End Using
        Return strRet
    End Function
    Private Function GetClientUpdate(clientID As String, connectionString As String) As String
        Dim strRet As String = String.Empty
        Using conn As SqlConnection = New SqlConnection(connectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand("Select Top 1 MercuryCutoffDate from mrtDataMart1.mrtMeritSharedDatamart.[MERIT_MATCH].[MercuryProjects] where ClientID = " & clientID.ToString &
            " Order by MercuryProjectID Desc")
                cmd.Connection = conn
                strRet = cmd.ExecuteScalar()
            End Using
        End Using
        Return strRet
    End Function
End Module
