Imports System.Data : Imports System.Data.SqlServerCe

Public Class SQLControl
    ' SQL CONNECTIONS
    Private DBcon As New SqlCeConnection("Data Source=MartinEnergy.sdf;")
    Private DBcmd As SqlCeCommand
    ' SQL DATA
    Public DBDA As SqlCeDataAdapter
    Public DBDS As DataSet
    ' QUERY PARAMETERS
    Public Params As New List(Of SqlCeParameter)
    ' QUERY STATISTICS
    Public RecordCount As Integer
    Public Exception As String

    Public Function HasConnection() As Boolean
        Exception = Nothing
        Try
            DBcon.Open() : DBcon.Close() : Return True
        Catch ex As Exception
            Exception = ex.Message
        End Try
        If DBcon.State = ConnectionState.Open Then DBcon.Close()
        Return False
    End Function

    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New SqlCeParameter(Name, Value)
        Params.Add(NewParam)
    End Sub

    Public Sub ExecQuery(Query As String)
        'RESET RECORD COUNT FOR NEW QUERIES
        RecordCount = 0
        Exception = Nothing
        Try
            DBcon.Open()

            DBcmd = New SqlCeCommand(Query, DBcon) ' CREATE SQL COMMAND
            Params.ForEach(Sub(p) DBcmd.Parameters.Add(p)) ' LOAD PARAMETERS INTO SQL COMMAND
            Params.Clear() ' CLEAR PARAMETER LIST
            ' EXECUTE COMMAND AND FILL DATASET
            DBDS = New DataSet
            DBDA = New SqlCeDataAdapter(DBcmd)
            RecordCount = DBDA.Fill(DBDS)

            DBcon.Close()
        Catch ex As Exception ' CAPTURE ERRORS
            Exception = ex.Message
        End Try
        'MAKE SURE THE CONNECTION IS CLOSED
        If DBcon.State = ConnectionState.Open Then DBcon.Close()
    End Sub

End Class
