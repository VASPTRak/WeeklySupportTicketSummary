Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports Azure
Imports log4net
Imports log4net.Config
Imports Mysqlx.XDevAPI
Imports NPOI.HSSF.Record
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.Formula.Functions
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel
Imports SixLabors.Fonts
Imports Ubiety.Dns.Core

Module Module1
    Private ReadOnly log As ILog = LogManager.GetLogger(GetType(Module1))

    Sub Main()
        XmlConfigurator.Configure()
        log.Info("Execution started")
        Try
            Dim body As String = String.Empty
            Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForSupportDataToGroupAdminEmailTemplate"))
                body = sr.ReadToEnd()
            End Using
            Dim filePath As String = ConfigurationManager.AppSettings("PathForSaveOpenTicketReport").ToString()
            log.Debug("Template Path: " + ConfigurationManager.AppSettings("PathForSupportDataToGroupAdminEmailTemplate") + " File folder: " + ConfigurationManager.AppSettings("PathForSaveOpenTicketReport").ToString())
            StartProcessing()
        Catch ex As Exception
            log.Error("Error occurred in Main. ex is :" & ex.Message)
        End Try
        log.Info("Execution stopped")
    End Sub

    Private Sub StartProcessing()
        Try
            log.Info("In StartProcessing")

            'Check and get Group Admin and support data
            Dim day As String = ConfigurationManager.AppSettings("Day").ToString().ToLower()
            Dim dayOfWeek As String = Date.Today.ToString("dddd").ToString().ToLower()
            Dim time As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("Time").ToString())


            Dim timeUtc = DateTime.UtcNow
            Dim easternZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")
            Dim easternTime As DateTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, easternZone)

            Dim timeOfDay As Integer = easternTime.TimeOfDay.Hours

            log.Debug("Day of week: " & dayOfWeek & " Eastern Time of the Day: " & timeOfDay)

            If day = dayOfWeek Then
                If time <= timeOfDay Then
                    Dim OBJMasterBAL = New MasterBAL()
                    Dim checkDataSendFlag As Boolean = OBJMasterBAL.CheckOrSetDataSendFlag(0)
                    If checkDataSendFlag Then

                        Dim dsGroupAdmins As DataSet = OBJMasterBAL.GetGroupAdminsHavingAccessToTicket()

                        If dsGroupAdmins IsNot Nothing Then
                            If dsGroupAdmins.Tables(0) IsNot Nothing Then
                                If dsGroupAdmins.Tables(0).Rows.Count > 0 Then
                                    For i = 0 To dsGroupAdmins.Tables(0).Rows.Count - 1
                                        Try
                                            OBJMasterBAL = New MasterBAL()
                                            Dim dsSuppData As DataSet = OBJMasterBAL.GetSupportItemsByGroupAdmin(dsGroupAdmins.Tables(0).Rows(i)(0))
                                            If dsSuppData IsNot Nothing Then
                                                If dsSuppData.Tables(0) IsNot Nothing Then
                                                    If dsSuppData.Tables(0).Rows.Count > 0 Then
                                                        ExcelProcessing(dsSuppData.Tables(0), dsGroupAdmins.Tables(0).Rows(i)(1))
                                                        log.Debug("Done for: " & dsGroupAdmins.Tables(0).Rows(i)(0).ToString())
                                                    Else
                                                        log.Debug("No data found: " & dsGroupAdmins.Tables(0).Rows(i)(0).ToString())
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Exception
                                            log.Error("Error occurred in StartProcessing for loop. ex is :" & ex.Message)
                                        End Try
                                    Next
                                Else
                                    log.Debug("No data(admin) found.")
                                End If
                            End If
                        End If

                        ' set flag for Sunday 5 PM EDT 
                        OBJMasterBAL.CheckOrSetDataSendFlag(1)
                    End If
                End If
            End If
        Catch ex As Exception
            log.Error("Error occurred in StartProcessing. ex is :" & ex.Message)
        End Try
    End Sub

    Private Sub ExcelProcessing(dtData As DataTable, EmailId As String)
        Try
            log.Info("In ExcelProcessing")

            '========== Main Support Data ======================
            Dim dtTransactions As DataTable = dtData
            Dim dtFinal As DataTable = New DataTable()

            Dim columns As String = "Company,Issue Type,Replacement Part Ordered,Status,Date,Created By"
            Dim columnsOfArray() As String = columns.Split(",")

            For l As Integer = 0 To columnsOfArray.Count - 1
                dtFinal.Columns.Add(columnsOfArray(l).ToString(), System.Type.[GetType]("System.String"))
            Next

            Dim drColumnNameNew As DataRow = dtFinal.NewRow()

            Dim columnNameList As List(Of String) = New List(Of String)
            For Each column As DataColumn In dtFinal.Columns
                columnNameList.Add(column.ColumnName)
            Next

            drColumnNameNew.ItemArray = columnNameList.ToArray()
            dtFinal.Rows.Add(drColumnNameNew)

            For Each dr As DataRow In dtTransactions.Rows
                Try
                    Dim drNew As DataRow = dtFinal.NewRow()
                    drNew("Company") = dr("Company").ToString()
                    drNew("Issue Type") = dr("IssueTypeText").ToString()
                    drNew("Replacement Part Ordered") = dr("ReplacementStuff").ToString()
                    drNew("Status") = dr("StatusText").ToString()
                    drNew("Date") = dr("SupportDate").ToString()
                    drNew("Created By") = dr("CaseOpenedBy").ToString()

                    dtFinal.Rows.Add(drNew)
                Catch ex As Exception
                    log.Error("Error occurred in ExcelProcessing for loop. ex is :" & ex.Message)
                End Try
            Next
            '==================================================================

            Dim FileName = WriteExcelWithNPOI("xlsx", dtFinal)

            SendEmail(FileName, EmailId) '

        Catch ex As Exception
            log.Error("Error occurred in ExcelProcessing. ex is :" & ex.Message)
        End Try
    End Sub

    Public Function WriteExcelWithNPOI(ByVal extension As String, ByVal dtData As DataTable) As String
        Dim fullPath = ""
        Try

            Dim workbook As IWorkbook

            If extension = "xlsx" Then
                workbook = New XSSFWorkbook()
            ElseIf extension = "xls" Then
                workbook = New HSSFWorkbook()
            Else
                Throw New Exception("This format is not supported")
            End If

            Dim rowCounter As Integer = 0
            Dim sheet1 As ISheet = workbook.CreateSheet("Sheet 1")

            Dim boldFont As XSSFFont = workbook.CreateFont()
            boldFont.IsBold = True

            ' create bordered cell style
            Dim borderedHeaderCellStyle As XSSFCellStyle = workbook.CreateCellStyle()
            borderedHeaderCellStyle.SetFont(boldFont)
            borderedHeaderCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium

            Dim borderedCellStyle As XSSFCellStyle = workbook.CreateCellStyle()
            borderedCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin

            Dim customStyle As XSSFCellStyle = workbook.CreateCellStyle()
            customStyle.SetFont(boldFont)

            Dim rowHeader As IRow = sheet1.CreateRow(0)
            Dim cellHeader As ICell = rowHeader.CreateCell(0)
            cellHeader.SetCellValue("Period Covered")
            cellHeader.CellStyle = customStyle
            cellHeader = rowHeader.CreateCell(1)
            Dim weekdate As String = ""
            weekdate = Date.Today.ToString("MMM dd,yyyy") & " - " & DateTime.Now.AddDays(-7).ToString("MMM dd,yyyy")
            cellHeader.SetCellValue(weekdate)
            cellHeader.CellStyle = customStyle

            '========== Main Transaction Data ======================
            For i As Integer = 0 To dtData.Rows.Count - 1
                rowCounter = rowCounter + 1
                Dim row As IRow = sheet1.CreateRow(i + 1) ' One row already added before so to balance added 1+
                For j As Integer = 0 To dtData.Columns.Count - 1
                    Dim cell As ICell = row.CreateCell(j)
                    Dim columnName As String = dtData.Columns(j).ToString()
                    cell.SetCellValue(dtData.Rows(i)(columnName).ToString())
                    If i = 0 Then
                        cell.CellStyle = borderedHeaderCellStyle
                    Else
                        cell.CellStyle = borderedCellStyle
                    End If
                Next
            Next

            '==================================================================

            Dim filePath As String = ConfigurationManager.AppSettings("PathForSaveOpenTicketReport").ToString()
            fullPath = filePath & DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss_fff") & ".xls"


            Using exportData = New MemoryStream()
                workbook.Write(exportData)
                Try
                    If (File.Exists(fullPath)) Then
                        'File.Delete(fullPath)
                        System.IO.File.Delete(fullPath)
                    End If
                Catch ex As Exception
                    log.Info("In WriteExcelWithNPOI => exception in delete file. filename: " & fullPath & "; exception is: " & ex.ToString())
                End Try


                log.Info("In WriteExcelWithNPOI => step 9")
                Dim bw As BinaryWriter = New BinaryWriter(File.Open(fullPath, FileMode.OpenOrCreate))

                bw.Write(exportData.ToArray())

                log.Info("In WriteExcelWithNPOI => step 10")
                bw.Close()
                workbook.Close()

            End Using

        Catch ex As Exception
            If Not (ex.Message.Contains("Thread was being aborted") = True) Then
                log.Error("Error occurred in WriteExcelWithNPOI. Exception is :" + ex.Message)
            End If
        End Try
        Return fullPath
    End Function

    Private Sub SendEmail(FileName As String, EmailId As String)
        Try

            Dim body As String = String.Empty
            Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForSupportDataToGroupAdminEmailTemplate"))
                body = sr.ReadToEnd()
            End Using
            '------------------

            body = body.Replace("owneremail", EmailId)

            Try
                body = body.Replace("ImageSign", "<img src=""https://www.fluidsecure.net/Content/Images/FluidSECURELogo.png"" style=""width:200px""/>")
                body = body.Replace("SupportTeamName", "FluidSecure Support Team")
                body = body.Replace("supportemail", "support@fluidsecure.com")
                body = body.Replace("SupportPhoneNumber", "1-850-878-4585")
                body = body.Replace("SupportLine1", "Press ""0"" During Normal Business Hours:  Monday - Friday 8:00am - 5:00pm (EST)")
                body = body.Replace("SupportLine2", "Press ""7"" After Normal Business Hours")
                body = body.Replace("websiteURLHREF", "https://www.fluidsecure.com")
                body = body.Replace("webisteURL", "www.fluidsecure.com")
            Catch ex As Exception
                body = body.Replace("ImageSign", "")
            End Try

            Dim mailClient As New SmtpClient(ConfigurationManager.AppSettings("smtpServer"))
            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("emailAccount"), ConfigurationManager.AppSettings("emailPassword"))
            mailClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings("smtpPort"))

            Dim messageSend As New MailMessage()
            messageSend.Body = body
            messageSend.IsBodyHtml = True
            messageSend.Subject = "***Group Admin Weekly Summary Ticket Report.***"
            messageSend.From = New MailAddress(ConfigurationManager.AppSettings("FromEmail"))

            If FileExists(FileName) Then
                Dim attach As Attachment = New Attachment(FileName)
                attach.Name = ConfigurationManager.AppSettings("FileName").ToString()
                messageSend.Attachments.Add(attach)
            End If

            mailClient.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings("EnableSsl"))

            If EmailId <> "" Then
                messageSend.To.Add(EmailId.Trim()) '
                mailClient.Send(messageSend)
                log.Info("Email send to: " + EmailId)
                messageSend.Attachments.Clear()
                messageSend.To.Remove(New MailAddress(EmailId.Trim())) '
                Try
                    System.IO.File.Delete(FileName)
                Catch ex As Exception
                    log.Error("When deleting file after email send : " + ex.ToString())
                End Try
            End If


        Catch ex As Exception
            log.Debug("Exception occurred in while sending email to " & EmailId & " . ex is :" & ex.ToString())
        End Try
    End Sub

    Private Function FileExists(ByVal FileFullPath As String) _
     As Boolean
        Try
            If FileFullPath = "" Then Return False

            Dim f As New IO.FileInfo(FileFullPath)
            Return f.Exists

        Catch ex As Exception
            log.Error("Exception occurred in FileExists. ex is :" & ex.ToString())
        End Try
    End Function
End Module
