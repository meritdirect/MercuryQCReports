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
        Dim outputFile As String = String.Empty
        Try
            Select Case strReportType
                Case "Element"
                    outputFile = runElementReport(clientID)
                Case "Mailed"
                    outputFile = runCampaignMatchReport(clientID, 2)
                Case "CampaignMatch"
                    outputFile = runCampaignMatchReport(clientID, 1)
                Case "OrderSummary"
                    outputFile = runOrderSummaryReport(clientID)
                Case "MatchLevel"
                    outputFile = runMatchLevelReport(clientID)
                Case "MailedData"
                    outputFile = runMailedDataReport(clientID)
                Case Else

            End Select
        Catch ex As Exception

            LogError(ex.ToString)
            Console.Write(ex.ToString)
            System.Environment.Exit(1)
        End Try
        Console.WriteLine(outputFile)
        System.Environment.Exit(0)

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
                adapter.SelectCommand.CommandTimeout = 800

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
    Private Function runElementReport(clientID As Integer) As String
        Dim ret As String = String.Empty
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "ElementQC_Template.xlsx")

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
                ret = fiOut.Name
            End Using
            Return ret

    End Function
    Private Function runMailedReport(clientID As Integer) As String
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        Dim ret As String = String.Empty
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "MailedQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets("PivotData")
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            ' cmd.CommandText = "Exec MercuryAdmin.dbo.p_MailedQCReport " & clientID.ToString
            cmd.CommandText = "declare @ClientID int
SET NOCOUNT ON
Declare @Cipher varchar(100), @SQL varchar(Max)

set @ClientID = #ClientID

Select @Cipher =Cipher  from mrtDatamart1.[mrtMeritSharedDatamart].[MERIT_MATCH].[Clients] with (NOLOCK) where ClientID = @ClientID


declare @MercuryProjectID int
Set @MercuryProjectID = (
              Select  max(MercuryProjectID)
              From    mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.MercuryProjects with (NOLOCK)
              Where   ClientID = @ClientID
              And     MercuryProjectStatus in ('A','C')
         )

Select DatasetID, CampaignID into ##DS from 
mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.Datasets D with (nolock)
                         Inner Join mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.MercuryProjects MP with (NOLOCK)
                              On D.ClientID = MP.ClientID
                 Where   D.ClientID          = @ClientID
                And     MP.MercuryProjectID =@MercuryProjectID
                And     D.DropDate Between MP.MercuryAnalysisStartDate And MercuryCutoffDate

Select * into ##Mailed from tigerwood.MeritMatch.MERIT_MATCH.Mailed with (NOLOCK) where DatasetID in (Select DataSetID from ##DS)


Select  SelectID       = L.ListID,
             SelectName = I.ItemName + ' (Select ' + RTrim(Cast(L.ListID As VarChar)) + ')'
     into ##Selects
     From    mrtDD.mrtMyMDB.dbo.Lists L With (NoLock)
             Inner Join mrtDD.mrtMyMDB.dbo.Items I With (NoLock)
                  On L.ItemID = I.ItemID
	where L.ListID in (Select distinct(SelectID) from ##Mailed)

Select CampaignID, CampaignName, DropDateFirst into ##Campaigns from mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.Campaigns C with (nolock) where CampaignID in (Select distinct(CampaignID) from ##Mailed)

Set @SQL = '
SELECT C.CampaignID, C.CampaignName,  I.SelectID,   Cast(C.DropDateFirst as Date) as DropDate,
IsNull(SelectName,''UNASSIGNED'') as SelectName, 
Count(distinct I.MailedID) as Mailed, count(Distinct R.OrderNo) as Response, sum(R.OrderAmount) as ResponseDollars from ##Mailed I 
JOIN ##DS DS on I.DatasetID = DS.DatasetID
JOIN ##Selects S on I.SelectID = S.SelectID
JOIN ##Campaigns C with (nolock) on DS.CampaignID = C.CampaignID
LEFT OUTER JOIN tigerwood.MeritMatch.' + @Cipher + '.Responders R with (nolock) on I.MailedID = R.MailedID
GROUP BY C.CampaignID, C.CampaignName, Cast(C.DropDateFirst as Date), I.SelectID,  IsNull(SelectName,''UNASSIGNED'')'
print @SQL
EXEC (@SQL)

DROP TABLE IF EXISTS ##DS
DROP TABLE IF EXISTS ##Mailed
DROP TABLE IF EXISTS ##Selects
DROP TABLE IF EXISTS ##Campaigns
"
            cmd.CommandText = cmd.CommandText.Replace("#ClientID", clientID.ToString)
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionStringTigerwood).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case dt.Columns(nCol).ColumnName
                            Case "Mailed"
                                wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                            Case "DropDate"
                                wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "yyyy-mm-dd"
                                wsQC.SetValue(nRow, nCol + 1, CDate(dr.Item(nCol).ToString))
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
            ret = fiOut.Name
        End Using
        Return ret

    End Function
    Private Function runCampaignMatchReport(clientID As Integer, nType As Integer) As String
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        Dim fi As FileInfo
        Dim ret As String = String.Empty
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Select Case nType
            Case 1
                fi = New FileInfo(My.Settings.TemplateDir & "CampaignMatchQC_Template.xlsx")
            Case 2
                fi = New FileInfo(My.Settings.TemplateDir & "MailedQC_Template.xlsx")
            Case Else
                Return ""
        End Select


        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets("PivotData")
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            ' cmd.CommandText = "Exec MercuryAdmin.dbo.p_CampaignMatchQCReport " & clientID.ToString
            cmd.CommandText = "Declare @Cipher varchar(100), @SQL varchar(Max), @ClientID int

set @ClientID = #ClientID

Select @Cipher =Cipher  from mrtDatamart1.[mrtMeritSharedDatamart].[MERIT_MATCH].[Clients] with (NOLOCK) where ClientID = @ClientID


declare @MercuryProjectID int
Set @MercuryProjectID = (
              Select  max(MercuryProjectID)
              From    mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.MercuryProjects with (NOLOCK)
              Where   ClientID = @ClientID
              And     MercuryProjectStatus in ('A','C')
         )

Select DatasetID, CampaignID into ##DS from 
mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.Datasets D with (nolock)
                         Inner Join mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.MercuryProjects MP with (NOLOCK)
                              On D.ClientID = MP.ClientID
                 Where   D.ClientID          = @ClientID
                And     MP.MercuryProjectID =@MercuryProjectID
                And     D.DropDate Between MP.MercuryAnalysisStartDate And MercuryCutoffDate


Select @SQL = 'Select * into ##Mailed from tigerwood.MERITMATCHWork.' + @Cipher + '.Mailed with (NOLOCK) where DatasetID in (Select DataSetID from ##DS)'
EXEC (@SQL)

Select  SelectID       = L.ListID,
             SelectName = I.ItemName + ' (Select ' + RTrim(Cast(L.ListID As VarChar)) + ')'
     into ##Selects
     From    mrtDD.mrtMyMDB.dbo.Lists L With (NoLock)
             Inner Join mrtDD.mrtMyMDB.dbo.Items I With (NoLock)
                  On L.ItemID = I.ItemID
	where L.ListID in (Select distinct(SelectID) from ##Mailed)

Select CampaignID, CampaignName, DropDateFirst into ##Campaigns from mrtDatamart1.mrtMeritSharedDatamart.MERIT_MATCH.Campaigns C with (nolock) where CampaignID in (Select distinct(CampaignID) from ##Mailed)

Set @SQL = '
SELECT C.CampaignID, C.CampaignName,  I.SelectID,   Cast(C.DropDateFirst as Date) as DropDate,
IsNull(SelectName,''UNASSIGNED'') as SelectName, 
Count(distinct I.MailedID) as Mailed, count(Distinct R.OrderNo) as Response, sum(R.OrderAmount) as ResponseDollars from ##Mailed I 
JOIN ##DS DS on I.DatasetID = DS.DatasetID
JOIN ##Selects S on I.SelectID = S.SelectID
JOIN ##Campaigns C with (nolock) on DS.CampaignID = C.CampaignID
LEFT OUTER JOIN tigerwood.MeritMatch.' + @Cipher + '.Responders R with (nolock) on I.MailedID = R.MailedID
GROUP BY C.CampaignID, C.CampaignName, Cast(C.DropDateFirst as Date), I.SelectID,  IsNull(SelectName,''UNASSIGNED'')'
print @SQL
EXEC (@SQL)

DROP TABLE IF EXISTS ##DS
DROP TABLE IF EXISTS ##Mailed
DROP TABLE IF EXISTS ##Selects
DROP TABLE IF EXISTS ##Campaigns
"
            cmd.CommandText = cmd.CommandText.Replace("#ClientID", clientID.ToString)
            Dim dDollars As Double = 0.0
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionStringTigerwood).Tables(0)
                For Each dr As DataRow In dt.Rows
                    For nCol As Integer = 0 To dt.Columns.Count - 1
                        Select Case nType
                            Case 1
                                Select Case dt.Columns(nCol).ColumnName
                                    Case "Mailed", "Response"
                                        wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                                    Case "ResponseDollars"
                                        If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                            wsQC.SetValue(nRow, nCol + 1, dDollars)
                                        Else
                                            wsQC.SetValue(nRow, nCol + 1, 0.0)
                                        End If
                                    Case "DropDate"
                                        wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "yyyy-mm-dd"
                                        wsQC.SetValue(nRow, nCol + 1, CDate(dr.Item(nCol).ToString))
                                    Case Else
                                        wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                                End Select
                            Case 2
                                Select Case dt.Columns(nCol).ColumnName
                                    Case "Mailed"
                                        wsQC.SetValue(nRow, nCol + 1, CInt(dr.Item(nCol)))
                                    Case "DropDate"
                                        wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "yyyy-mm-dd"
                                        wsQC.SetValue(nRow, nCol + 1, CDate(dr.Item(nCol).ToString))
                                    Case Else
                                        wsQC.SetValue(nRow, nCol + 1, dr.Item(nCol).ToString)
                                End Select
                        End Select
                    Next
                    nRow += 1
                Next


            End Using
            Dim fiOut As FileInfo
            Select Case nType
                Case 1
                    fiOut = New FileInfo(reportDir & "CampaignMatchQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
                Case 2
                    fiOut = New FileInfo(reportDir & "MailedQCReport_" & clientName & "_" & clientUpdate & ".xlsx")
            End Select

            pck.SaveAs(fiOut)
            ret = fiOut.Name
        End Using
        Return ret

    End Function
    Private Function runOrderSummaryReport(clientID As Integer) As String
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        Dim ret As String = String.Empty
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
            ret = fiOut.Name
        End Using
        Return ret

    End Function
    Private Function runMatchLevelReport(clientID As Integer) As String
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        Dim ret As String = String.Empty
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
            ret = fiOut.Name
        End Using
        Return ret

    End Function
    Private Function runMailedDataReport(clientID As Integer) As String
        Dim clientName As String = GetClientName(clientID, My.Settings.ConnectionString)
        Dim clientUpdate As String = GetClientUpdate(clientID, My.Settings.ConnectionString)
        Dim ret As String = String.Empty
        clientUpdate = clientUpdate.Replace("\", "_").Replace("/", "_")
        Dim fi As FileInfo = New FileInfo(My.Settings.TemplateDir & "MailedDataQC_Template.xlsx")
        Using pck As ExcelPackage = New ExcelPackage(fi)
            Dim nRow As Integer = 2
            Dim wsQC As ExcelWorksheet = pck.Workbook.Worksheets(1)
            wsQC.DeleteRow(2, wsQC.Dimension.End.Row)
            Dim cmd As New SqlClient.SqlCommand
            cmd.CommandText = "Exec MercuryAdmin.dbo.p_MailedDataQCReport " & clientID.ToString
            Dim dDollars As Double = 0.0, iInt As Integer = 0
            Using dt As DataTable = ExecuteCMD(cmd, My.Settings.ConnectionString).Tables(0)
                For Each dr As DataRow In dt.Rows
                    If Not dr.Item(0).ToString.Contains("Sub-Total") Then 'skip SubTotals
                        For nCol As Integer = 0 To dt.Columns.Count - 2 'skip total sort
                            Select Case dt.Columns(nCol).ColumnName
                                Case "Mailed", "Response", "Auxiliary Orders"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Integer.TryParse(dr.Item(nCol), iInt) Then
                                        wsQC.SetValue(nRow, nCol + 1, iInt)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0)
                                    End If
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "#,##0"
                                Case "$/M Index", "RR Index", "AOV Index", "Aux$/Resp Index", "Aux Order/Resp Index"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                        wsQC.SetValue(nRow, nCol + 1, dDollars / 100)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0.0)
                                    End If
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "0.00%"
                                Case "Response Rate"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                        wsQC.SetValue(nRow, nCol + 1, dDollars / 100)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0.0)
                                    End If
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "0.00%"
                                Case "Response Dollars", "$/M", "AOV", "Auxiliary Dollars", "Aux$/Response", "GM", "AdCost", "OrderProcessing", "Contribution"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                        wsQC.SetValue(nRow, nCol + 1, dDollars)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0.0)
                                    End If
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "\$#,##0"
                                Case "$/Piece", "Aux Orders/Response", "Con/Order"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                        wsQC.SetValue(nRow, nCol + 1, dDollars)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0.0)
                                    End If
                                    wsQC.Cells(nRow, nCol + 1).Style.Numberformat.Format = "\$#,###.00"
                                Case "Aux Orders/Response"
                                    If Not IsDBNull(dr.Item(nCol)) AndAlso Double.TryParse(dr.Item(nCol), dDollars) Then
                                        wsQC.SetValue(nRow, nCol + 1, dDollars)
                                    Else
                                        wsQC.SetValue(nRow, nCol + 1, 0.0)
                                    End If
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
            ret = fiOut.Name
        End Using
        Return ret

    End Function
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
