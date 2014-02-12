#Region "Compile Options"
Option Strict On
Option Explicit On
#End Region

#Region "Import Namespaces"
Imports MosaicDBAccess
#End Region

Module ModuleFunctions
#Region "Public variables"
    Public mosaicDll As New MosaicDBAccess.MosaicDBAccessClass

    Public Const executionStatusFailedAbort As Integer = 0 '"Failed/Stop Run" - do not change this
    Public Const executionStatusFailed As Integer = 1 '"Failed" - do not change this 
    Public Const executionStatusPassed As Integer = 2 '"Passed" - do not change this 
    Public Const executionStatusNotCompleted As Integer = 3 '"Not completed" - do not change this 
    Public Const executionStatusNotRun As Integer = 4 '"Not run" - do not change this 
    Public Const executionStatusSkipped As Integer = 5 '"Skipped" - do not change this 
    Public Const executionLog As String = "Logs\Execution.log" 'a file written by the Selenium Actions dll
    Public Const senderMCPAction As String = "MCP Action" 'the sender text used by MosaicDBAccess.Logger.  Using this sender will cause "Actual Result|" to be recorded into the db.
    Public Const senderMCP As String = "MCP" 'the sender text used by MosaicDBAccess.Logger
#End Region

#Region "Public Functions"
#End Region

#Region "Private Functions"
#End Region
End Module
