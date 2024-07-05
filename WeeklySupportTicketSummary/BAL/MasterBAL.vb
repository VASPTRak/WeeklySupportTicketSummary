Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config
Public Class MasterBAL
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(MasterBAL))

    Shared Sub New()
        XmlConfigurator.Configure()
    End Sub

    Public Function CheckOrSetDataSendFlag(flag As Boolean) As Boolean
        Dim dal = New GeneralizedDAL()
        Dim result As Boolean = False
        Try

            Dim parcollection As SqlParameter() = New SqlParameter(0) {}

            parcollection(0) = New SqlParameter("@flag", SqlDbType.Bit)
            parcollection(0).Direction = ParameterDirection.Input
            parcollection(0).Value = flag

            result = dal.ExecuteStoredProcedureGetBoolean("usp_tt_Support_CheckOrSetWeeklyDataSendFlagToGroupAdmin", parcollection)

            Return result

        Catch ex As Exception

            log.Error("Error occurred in CheckOrSetDataSendFlag Exception is :" + ex.Message)
            Return False
        Finally

        End Try
    End Function

    Public Function GetGroupAdminsHavingAccessToTicket() As DataSet
        Dim dal = New GeneralizedDAL()
        Dim ds As DataSet = New DataSet()
        Try

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_Support_GetGroupAdminsHavingAccessToTicket")

            Return ds

        Catch ex As Exception

            log.Error("Error occurred in GetGroupAdminsHavingAccessToTicket Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function GetSupportItemsByGroupAdmin(Personid As Integer) As DataSet
        Dim dal = New GeneralizedDAL()
        Dim ds As DataSet = New DataSet()
        Try

            Dim parcollection As SqlParameter() = New SqlParameter(0) {}

            parcollection(0) = New SqlParameter("@Personid", SqlDbType.Int)
            parcollection(0).Direction = ParameterDirection.Input
            parcollection(0).Value = Personid

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_Support_GetSupportItemsByGroupAdmin", parcollection)

            Return ds

        Catch ex As Exception

            log.Error("Error occurred in GetSupportItemsByGroupAdmin Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function
End Class
