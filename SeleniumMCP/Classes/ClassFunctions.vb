#Region "Compile Options"
Option Strict On
Option Explicit On
#End Region

#Region "Import Namespaces"
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.IE
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.UI
Imports Selenium.DefaultSelenium
Imports System.Collections.ObjectModel
Imports SNAGITLib
Imports MosaicDBAccess
Imports Microsoft.Office.Interop
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data

'Imports OpenQA.Selenium.Remote.RemoteWebDriver
#End Region

Public Class ClassFunctions
    Dim WaitForTime As Integer = 20
    Public Class AutomationObject

        Private _Object As Object = Nothing
        Private _Element As IWebElement = Nothing

        Public Property aObject() As Object
            Get
                Return _Object
            End Get
            Set(ByVal value As Object)
                _Object = value
            End Set
        End Property
        Public Property aElement() As IWebElement
            Get
                Return _Element
            End Get
            Set(ByVal value As IWebElement)
                _Element = value
            End Set
        End Property



    End Class
    Public Class Locator
        Private _name As String
        Private _loctype As String
        Private _locText As String
        Private _PageName As String
        Private _LookupKey As String

        Public Property sName() As String
            Get
                Return _name.ToString
            End Get
            Set(ByVal value As String)
                _name = value
            End Set
        End Property

        Public Property sLocType() As String
            Get
                Return _loctype.ToString()
            End Get
            Set(ByVal value As String)
                _loctype = value
            End Set
        End Property

        Public Property sLocText() As String
            Get
                Return _locText.ToString()
            End Get
            Set(ByVal value As String)
                _locText = value
            End Set
        End Property

        Public Property sPageName() As String
            Get
                Return _PageName.ToString()
            End Get
            Set(ByVal value As String)
                _PageName = value
            End Set
        End Property

        Public Property s_LookupKey() As String
            Get
                Return _LookupKey.ToString
            End Get
            Set(ByVal value As String)
                _LookupKey = value
            End Set
        End Property


    End Class
#Region "Class Properties"
    Property LogPath() As String
        Get
            Return myLogPath
        End Get
        Set(ByVal value As String)
            myLogPath = value
        End Set
    End Property

    Property ExecutableHistoryRecordID() As Integer
        Get
            Return myExecutableHistoryRecordID
        End Get
        Set(ByVal value As Integer)
            myExecutableHistoryRecordID = value
        End Set
    End Property
#End Region
#Region "Private functions and procedures"
    ''' <summary>
    ''' Waits for a specified number of seconds for a WebElement to exist.
    ''' </summary>
    ''' <param name="delaySeconds">The time in seconds to wait for the element.</param>
    ''' <param name="locator">An OpenQA.Selenium.By class used to specify the WebElement in question.</param>
    ''' <returns>The specified element, or Nothing if the element can't be located within the time limit.</returns>
    ''' <remarks></remarks>
    Function WaitForElement(ByRef driver As IWebDriver, ByVal delaySeconds As Integer, ByVal locator As OpenQA.Selenium.By) As IWebElement
        Try
            Dim Wait As New WebDriverWait(driver, TimeSpan.FromSeconds(CDbl(delaySeconds)))
            WaitForElement = Wait.Until(Of IWebElement)(
                Function(d)
                    Return d.FindElement(locator)
                End Function
            )
        Catch ex As Exception
            ' Doesn't discriminate between a timeout and the Web Driver giving up from "no response." This just puts an upper
            ' limit on how long it can wait. delaySeconds can't be longer than the amount of time the Web Driver waits before
            ' it throws a "no response" exception.
            WaitForElement = Nothing
        End Try
    End Function
    Public Function WaitForElement(ByVal delaySeconds As Integer, ByVal locator As OpenQA.Selenium.By) As IWebElement
        Try
            Dim Wait As New WebDriverWait(driver, TimeSpan.FromSeconds(CDbl(delaySeconds)))
            WaitForElement = Wait.Until(Of IWebElement)(Function(d)
                                                            Return d.FindElement(locator)
                                                        End Function)
        Catch ex As Exception           ' Doesn't discriminate between a timeout and the Web Driver giving up from "no response." This just puts an upper
            WaitForElement = Nothing    ' limit on how long it can wait. delaySeconds can't be longer than the amount of time the Web Driver waits before
        End Try                         ' it throws a "no response" exception.
    End Function
    Public Function WaitForElements(ByRef driver As IWebDriver, ByVal delaySeconds As Integer, ByVal t As IWebElement, ByVal locator As OpenQA.Selenium.By) As ReadOnlyCollection(Of IWebElement)
        Try
            Dim Wait As New WebDriverWait(driver, TimeSpan.FromSeconds(CDbl(delaySeconds)))
            WaitForElements = Wait.Until(Of ReadOnlyCollection(Of IWebElement))(
                Function(d)
                    Return t.FindElements(locator)
                End Function
            )
        Catch ex As Exception
            ' Doesn't discriminate between a timeout and the Web Driver giving up from "no response." This just puts an upper
            ' limit on how long it can wait. delaySeconds can't be longer than the amount of time the Web Driver waits before
            ' it throws a "no response" exception.
            WaitForElements = Nothing
        End Try
    End Function
    ''' <summary>
    ''' Opens the (class scoped) connection to the SQL Server database for obtaining verification data.
    ''' </summary>
    ''' <param name="serverToUse"></param>
    ''' <returns>Integer value representing result of the call.</returns>
    ''' <remarks>Unlike the function below, does not constrain successive calls to the same DB server as the 1st call.</remarks>
    Private Function OpenSqlDB(ByVal serverToUse As String) As Integer
        Try
            If sqlConn IsNot Nothing AndAlso sqlConn.State = ConnectionState.Open Then
                Return dbOpen
            End If
            Dim connStr As String = "Server=" & serverToUse & ";Trusted_Connection=True;"
            If sqlConn Is Nothing Then
                sqlConn = New SqlConnection()
            End If
            sqlConn.ConnectionString = connStr
            sqlConn.Open()
            Return dbOpen
        Catch ex As Exception
            Return dbError
        End Try
    End Function
    ''' <summary>
    ''' Opens the (class scoped) connection to the SQL Server database for obtaining verification data.
    ''' </summary>
    ''' <param name="serverToUse"></param>
    ''' <returns>Integer value representing result of the call.</returns>
    ''' <remarks>After the first call to this function, all additional calls must specify the same DB server as the one used for the first call.</remarks>
    Private Function OpenConstantSqlDB(ByVal serverToUse As String) As Integer
        Static openServer As String = ""
        Try
            If sqlConn IsNot Nothing AndAlso sqlConn.State = ConnectionState.Open Then
                If sqlConn.DataSource.Trim() = serverToUse.Trim() Then
                    Return dbOpen
                Else
                    Return dbChanged
                End If
            End If
            If openServer <> "" AndAlso openServer <> serverToUse.Trim() Then
                Return dbChanged
            End If
            Dim connStr As String = "Server=" & serverToUse & ";Trusted_Connection=True;"
            If sqlConn Is Nothing Then
                sqlConn = New SqlConnection()
            End If
            sqlConn.ConnectionString = connStr
            sqlConn.Open()
            ' If this is the first successful call of the run, save the server string for subsequent calls.
            If openServer = "" Then
                openServer = serverToUse.Trim()
            End If
            Return dbOpen
        Catch ex As Exception
            Return dbError
        End Try
    End Function
    ''' <summary>
    ''' Closes the (class scoped) connection to the SQL Server database for obtaining verification data.
    ''' </summary>
    ''' <returns>Integer value representing result of the call.</returns>
    ''' <remarks>If the connection is already closed, or Nothing, returns dbClosed.</remarks>
    Private Function CloseSqlDB() As Integer
        If sqlConn Is Nothing OrElse sqlConn.State = ConnectionState.Closed Then
            Return dbClosed
        End If
        Try
            sqlConn.Close()
            sqlConn = Nothing
            Return dbClosed
        Catch ex As Exception
            Return dbError
        End Try
        sqlConn.Close()
        sqlConn = Nothing
    End Function
    ''' <summary>
    ''' Pads a numeric value stored in numberStr with leading zeros to make it totalLength long.
    ''' </summary>
    ''' <param name="numberStr">The String containing a positive numeric value.</param>
    ''' <param name="totalLength">The desired length of the returned String including any added leading zeros.</param>
    ''' <returns>The String containing the original numeric value with sufficient leading zeros added to make its length totalLength.</returns>
    ''' <remarks>If the length of numberStr is equal to or longer than totalLength, just returns numberStr unmodified.</remarks>
    Private Function PadWithLeadingZeros(ByVal numberStr As String, totalLength As Integer) As String
        Const excepMsg As String = "PadWithZeros(): Parameter ""numberStr"" cannot "
        If numberStr.Length < totalLength Then
            If IsNumeric(numberStr) Then
                If CLng(numberStr) < 0 Then
                    Throw New Exception(excepMsg & "contain a negative value.")
                End If
                PadWithLeadingZeros = CUInt(numberStr).ToString("D" & totalLength.ToString())
            Else
                Throw New Exception(excepMsg & "be converted into a numeric value.")
            End If
        Else
            PadWithLeadingZeros = numberStr
        End If
    End Function
    ''' <summary>
    ''' Removes commas and dollar sign from a string representing a monetary value.
    ''' </summary>
    ''' <param name="value">The string to be cleaned.</param>
    ''' <returns>String with the characters removed.</returns>
    ''' <remarks>Also removes leading and trailing white space. Input string values that contain parentheses are made negative.</remarks>
    Private Function CleanMonetaryString(ByVal value As String) As String
        Dim result As String = value.Replace(",", "").Replace("$", "").Trim()
        If result.Contains("(") AndAlso result.Contains(")") Then               ' It's a negative value, so...
            result = "-" & result.Replace("(", "").Replace(")", "").Trim()      ' ...remove the "()" and prepend a "negative sign" (hyphen).
        End If
        Return result
    End Function
    ''' <summary>
    ''' Takes a snapshot of the current screen with SnagIt.
    ''' </summary>
    ''' <param name="actionHistoryRecordID">The action history record ID</param>
    ''' <returns>An Integer signifying the outcome.</returns>
    ''' <remarks>The function also adds information to the log describing the outcome.</remarks>
    Private Function CaptureScreenshot(ByVal actionHistoryRecordID As Integer) As Integer
        Dim snagItDirectory As String = ""
        Dim snagItFileName As String = ""       ' File name w/o the extension.
        Dim snagItFileExtension As String = ""  ' Extension of the image file.
        Try
            ' Get the directory for the SnagIt path.
            snagItDirectory = My.Computer.FileSystem.CombinePath(
                My.Computer.FileSystem.GetParentPath(My.Computer.FileSystem.GetParentPath(myLogPath)), "Screenshots")

            ' Get the current date/time and use that as the SnagIt file name.
            snagItFileName = My.Computer.Clock.LocalTime.Year.ToString & _
                My.Computer.Clock.LocalTime.Month.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Day.ToString.PadLeft(2, "0"c) & _
                "_" & _
                My.Computer.Clock.LocalTime.Hour.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Minute.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Second.ToString.PadLeft(2, "0"c)

            ' Set SnagIt to take a desktop screenshot and store it as a file.
            snagIt.Input = SNAGITLib.snagImageInput.siiDesktop
            snagIt.Output = SNAGITLib.snagImageOutput.sioFile

            ' Set the screenshot properties and save location.
            snagIt.OutputImageFile.FileNamingMethod = SNAGITLib.snagOuputFileNamingMethod.sofnmFixed
            snagIt.OutputImageFile.Directory = snagItDirectory
            snagIt.OutputImageFile.Filename = snagItFileName
            snagIt.OutputImageFile.FileType = SNAGITLib.snagImageFileType.siftPNG
            Select Case snagIt.OutputImageFile.FileType
                Case snagImageFileType.siftPNG
                    snagItFileExtension = ".png"
                Case snagImageFileType.siftJPEG
                    snagItFileExtension = ".jpeg"
                Case snagImageFileType.siftGIF
                    snagItFileExtension = ".gif"
                Case snagImageFileType.siftBMP
                    snagItFileExtension = ".bmp"
                Case Else
                    snagItFileExtension = ""
            End Select

            ' Suppress snagit preview window and cursor.
            snagIt.EnablePreviewWindow = False
            snagIt.IncludeCursor = False

            ' Take a snapshot, check for SnagIt license.
            Try
                snagIt.Capture()
                mosaicDll.RSTAR_AddScreenshot(CLng(myExecutableHistoryRecordID), 5, CLng(actionHistoryRecordID), My.Computer.FileSystem.CombinePath(snagItDirectory, snagItFileName & snagItFileExtension), "") 'Use 5 to attach a screenshot to an Action in the history table.
                mosaicDll.Logger(senderMCPAction, "Screenshot|" & My.Computer.FileSystem.CombinePath(snagItDirectory, snagItFileName & snagItFileExtension), myLogPath, , actionHistoryRecordID)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Screenshot captured.", myLogPath, , actionHistoryRecordID)
            Catch ex As Exception
                If snagIt.LastError = SNAGITLib.snagError.serrSnagItExpired Then
                    CaptureScreenshot = snapExpired
                Else
                    CaptureScreenshot = snapFail
                End If
                mosaicDll.Logger(senderMCPAction, "Screenshot|Failure.", myLogPath, , actionHistoryRecordID)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Screenshot capture failed.", myLogPath, , actionHistoryRecordID)
                Exit Function
            End Try
            CaptureScreenshot = snapSuccess
        Catch ex As Exception
            CaptureScreenshot = snapUnknownException
        End Try
    End Function
#End Region
#Region "Private variables"
    Private driver As IWebDriver 'interface contract...generic object to invoke Selenium actions
    Private snagIt As SNAGITLib.IImageCapture2 = New SNAGITLib.ImageCaptureClass() 'using IImageCapture2 to get new functionality in SnagIt 8.1
    Private myLogPath As String
    Private myExecutableHistoryRecordID As Integer
    Private sqlConn As SqlConnection = Nothing

    Private Const dbError As Integer = -1                                           ' Return values for DB open/close functions.
    Private Const dbClosed As Integer = 0
    Private Const dbOpen As Integer = 1
    Private Const dbChanged As Integer = 2

    Private Const snapSuccess As Integer = 0                                        ' Return values for CaptureScreenshot()
    Private Const snapFail As Integer = -1
    Private Const snapExpired As Integer = -2
    Private Const snapUnknownException As Integer = -3

    Private Const nulFlg As String = "DBNull"
    Private Const exceptionErrorMsg As String = "Actual Result|Exception occurred. See log for details." ' Don't remove "Actual Result|" from beginning!


#End Region
#Region "Selenium Function Templates"

    ' These are templates for your test action functions.

    'Public Function ta_SeleniumActionTemplate(ByVal mcpParameters() As Object) As Object()
    '    Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
    '    Dim excelPath As String = CType(mcpParameters(0), String)
    '    Dim iteration As Integer = CType(mcpParameters(1), Integer)
    '    Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
    '    Try

    '        ''''''''''''''''''''
    '        ' Do something here.
    '        ''''''''''''''''''''

    '        ' ( **** Remove the next 2 (or 3) lines once the test action function is ready for testing. ****)
    '        returnValues(1) = "(* Dummy action *)"
    '        mosaicDll.Logger(senderMCPAction, "Actual Result|Dummy action", myLogPath, , actionHistoryRecordID)
    '        Return returnValues
    '    Catch ex As Exception
    '        returnValues(0) = executionStatusFailed
    '        returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
    '        mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
    '        mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
    '        Return returnValues
    '    End Try
    'End Function

    'Public Function ta_SeleniumActionTemplate_with_DB_Access(ByVal mcpParameters() As Object) As Object()
    '    Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
    '    Dim excelPath As String = CType(mcpParameters(0), String)
    '    Dim iteration As Integer = CType(mcpParameters(1), Integer)
    '    Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
    '    Try
    '        ' ( **** Uncomment the section below when your ready to actually test this function. **** )
    '        '' Open the DB.
    '        'Dim dbServer As String =  (Name of database here.)                                    ' (FIX THIS LINE BY DEFINING THE DATABASE TO BE ACCESSED!)
    '        'Select Case OpenSqlDB(dbServer)
    '        '    Case dbError
    '        '        returnValues(0) = executionStatusFailed
    '        '        returnValues(1) = "Unable to open '" & dbServer & "' database."
    '        '        mosaicDll.Logger(senderMCPAction, "Actual Result|Unable to open '" & dbServer & "' database.", myLogPath, , actionHistoryRecordID)
    '        '    Case dbChanged
    '        '        returnValues(0) = executionStatusFailed
    '        '        returnValues(1) = "Data profile specifies a changed database URL."
    '        '        mosaicDll.Logger(senderMCPAction, "Actual Result|Data profile specifies a changed database URL.", myLogPath, , actionHistoryRecordID)
    '        'End Select
    '        'If CInt(returnValues(0)) <> executionStatusPassed Then
    '        '    Exit Try
    '        'End If

    '        ''''''''''''''''''''
    '        ' Do something here.
    '        ''''''''''''''''''''

    '        ' ( **** Remove the next 2 lines once the test action function is ready for testing. **** )
    '        returnValues(1) = "(* Dummy action *)"
    '        mosaicDll.Logger(senderMCPAction, "Actual Result|Dummy action", myLogPath, , actionHistoryRecordID)
    '    Catch ex As Exception
    '        returnValues(0) = executionStatusFailed
    '        returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
    '        mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
    '        mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
    '    Finally
    '        ' Close the DB.
    '        If CloseSqlDB() <> dbClosed Then
    '            returnValues(0) = executionStatusFailed
    '            returnValues(1) = "Error closing the database."
    '            mosaicDll.Logger(senderMCPAction, "Error|Error closing the database.", myLogPath, , actionHistoryRecordID)
    '        End If
    '        ta_SeleniumActionTemplate_with_DB_Access = returnValues
    '    End Try
    'End Function

#End Region

    Public Class CalendarParams

        Private _yeardec As String
        Private _monthdec As String
        Private _monthinc As String
        Private _yearinc As String
        Private _clean As String
        Private _today As String
        Private _CalButtonXPath As String
        Private _CalHeader As String
        Private _CurrentDay As Integer
        Private _CurrentMonth As Integer
        Private _CurrentYear As Integer
        Private _TargetDay As Integer
        Private _TargetMonth As Integer
        Private _TargetYear As Integer
        Private _jumpMonthsBy As Integer
        Private _increment As Boolean
        Private _dayXPathPart1 As String
        Private _dayXpathPart2 As String
        Private _dayZeroXPath As String
        Private _applyButtonXPath As String
        Private _TargetDate As DateTime
        Private _CurrentDate As DateTime
        Private _iMonthIncDecValue As Integer
        Private _iYearincDecValue As Integer
        Private _DayIndex As Integer


        Public Function IsPositive(number As Integer) As Boolean
            Return number > 0
        End Function
        Public Function IsNegative(number As Integer) As Boolean
            Return number < 0
        End Function
        Public Sub SetYearsAdndMonthsJump()

            _iMonthIncDecValue = CInt(DateDiff(DateInterval.Month, _CurrentDate, _TargetDate))
            _iYearincDecValue = CInt(DateDiff(DateInterval.Year, _CurrentDate, _TargetDate))

        End Sub
        Public Property DayIndex() As Integer
            Get
                Return _DayIndex
            End Get
            Set(ByVal value As Integer)
                _DayIndex = value
            End Set
        End Property
        Public Property MonthIncDecValue() As Integer
            Get
                Return _iMonthIncDecValue
            End Get
            Set(ByVal value As Integer)
                _iMonthIncDecValue = value
            End Set
        End Property
        Public Property iYearincDecValue() As Integer
            Get
                Return _iYearincDecValue
            End Get
            Set(ByVal value As Integer)
                _iYearincDecValue = value
            End Set
        End Property
        Public Property TargetDate() As DateTime
            Get
                Return _TargetDate
            End Get
            Set(ByVal value As DateTime)
                _TargetDate = value
            End Set
        End Property
        Public Property CurrentDate() As DateTime
            Get
                Return _CurrentDate
            End Get
            Set(ByVal value As DateTime)
                _CurrentDate = value
            End Set
        End Property
        Public Property applyButtonXPath() As String
            Get
                Return _applyButtonXPath
            End Get
            Set(ByVal value As String)
                _applyButtonXPath = value
            End Set
        End Property
        Public Property dayZeroXPath() As String
            Get
                Return _dayZeroXPath
            End Get
            Set(ByVal value As String)
                _dayZeroXPath = value
            End Set
        End Property
        Public Property dayXPathPart1() As String
            Get
                Return _dayXPathPart1
            End Get
            Set(ByVal value As String)
                _dayXPathPart1 = value
            End Set
        End Property
        Public Property dayXPathPart2() As String
            Get
                Return _dayXpathPart2
            End Get
            Set(ByVal value As String)
                _dayXpathPart2 = value
            End Set
        End Property
        Public Property TargetDay() As Integer
            Get
                Return _TargetDay
            End Get
            Set(ByVal value As Integer)
                _TargetDay = value
            End Set
        End Property
        Public Property TargetMonth() As Integer
            Get
                Return _TargetMonth
            End Get
            Set(ByVal value As Integer)
                _TargetMonth = value
            End Set
        End Property
        Public Property TargetYear() As Integer
            Get
                Return _TargetYear
            End Get
            Set(ByVal value As Integer)
                _TargetYear = value
            End Set
        End Property
        Public Property CurrentDay() As Integer
            Get
                Return _CurrentDay
            End Get
            Set(ByVal value As Integer)
                _CurrentDay = value
            End Set
        End Property
        Public Property CurrentMonth() As Integer
            Get
                Return _CurrentMonth
            End Get
            Set(ByVal value As Integer)
                _CurrentMonth = value
            End Set
        End Property
        Public Property CurrentYear() As Integer
            Get
                Return _CurrentYear
            End Get
            Set(ByVal value As Integer)
                _CurrentYear = value
            End Set
        End Property
        Public Property yeardec() As String
            Get
                Return _yeardec
            End Get
            Set(ByVal value As String)
                _yeardec = value
            End Set
        End Property
        Public Property monthdec() As String
            Get
                Return _monthdec
            End Get
            Set(ByVal value As String)
                _monthdec = value
            End Set
        End Property
        Public Property monthinc() As String
            Get
                Return _monthinc
            End Get
            Set(ByVal value As String)
                _monthinc = value
            End Set
        End Property
        Public Property yearinc() As String
            Get
                Return _yearinc
            End Get
            Set(ByVal value As String)
                _yearinc = value
            End Set
        End Property
        Public Property clean() As String
            Get
                Return _clean
            End Get
            Set(ByVal value As String)
                _clean = value
            End Set
        End Property
        Public Property ctoday() As String
            Get
                Return _today
            End Get
            Set(ByVal value As String)
                _today = value
            End Set
        End Property
        Public Property CalButtonXPath() As String
            Get
                Return _CalButtonXPath
            End Get
            Set(ByVal value As String)
                _CalButtonXPath = value
            End Set
        End Property

        Public Property CalHeader() As String
            Get
                Return _CalHeader
            End Get
            Set(ByVal value As String)
                _CalHeader = value
            End Set
        End Property

    End Class
    Public Class BrowseFileParms
        Private _RemovePath As String
        Private _AttachPath As String
        Private _FileToAttach As String
        Public Property RemovePath() As String
            Get
                Return _RemovePath
            End Get

            Set(ByVal value As String)
                _RemovePath = value
            End Set

        End Property
        Public Property AttachPath() As String
            Get
                Return _AttachPath
            End Get

            Set(ByVal value As String)
                _AttachPath = value
            End Set

        End Property
        Public Property FileToAttach() As String
            Get
                Return _FileToAttach
            End Get

            Set(ByVal value As String)
                _FileToAttach = value
            End Set

        End Property

    End Class
    Public Class TextAreaParms
        Private _TextToEnter As String
        Private _TextAreaXpath As String
        Public Property TextToEnter() As String
            Get
                Return _TextToEnter
            End Get

            Set(ByVal value As String)
                _TextToEnter = value
            End Set

        End Property
        Public Property TextAreaXpath() As String
            Get
                Return _TextAreaXpath
            End Get

            Set(ByVal value As String)
                _TextAreaXpath = value
            End Set

        End Property
    End Class
    Public Class TextBoxParms
        Private _TextStringToEnter As String
        Private _TextXPath As String
        Public Property TextStringToEnter() As String
            Get
                Return _TextStringToEnter
            End Get

            Set(ByVal value As String)
                _TextStringToEnter = value
            End Set

        End Property
        Public Property TextXpath() As String
            Get
                Return _TextXPath
            End Get

            Set(ByVal value As String)
                _TextXPath = value
            End Set

        End Property
    End Class
    Public Sub DoCalendar(ByVal sDateParm As String, ByVal CalendarRef As CalendarParams, ByRef returnValues() As Object, ByVal actionHistoryRecordID As Integer, Optional ByVal sCalendarName As String = "")
        'Dim calExpectedDrawDate As New CalendarParams
        'Dim properDate As DateTime = DateTime.Parse(sDateParm)
        Dim index As Integer = 0
        Const MAX_DAY_IN_A_MONTH As Integer = 31

        'open the calendar
        Dim Calendar_Found As Boolean = True
        Dim oCalender As Object = Nothing
        oCalender = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.CalButtonXPath))


        If Not oCalender Is Nothing Then
            Dim oCalendar_Element As IWebElement = DirectCast(oCalender, IWebElement)
            oCalendar_Element.Click()
            Dim oClean_Object As Object = Nothing
            Dim WaitForClean As Integer = 3
            oClean_Object = WaitForElement(driver, WaitForClean, By.XPath(CalendarRef.clean))
            If Not oClean_Object Is Nothing Then
                Dim oClean_Element As IWebElement = DirectCast(oClean_Object, IWebElement)
                oClean_Element.Click()
                oCalendar_Element.Click()
            End If
            mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the " & sCalendarName & " calendar button.", myLogPath, , actionHistoryRecordID)
        Else
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Could not click the " & sCalendarName & " calendar button."
            mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the " & sCalendarName & " calendar button.", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Calendar_Found = False
        End If

        If Calendar_Found = True Then
            'calculate the number of years and months to jump we only use months
            CalendarRef.SetYearsAdndMonthsJump()

            Dim yearObject As Object = Nothing
            Dim monthObject As Object = Nothing
            Dim yearElement As IWebElement = Nothing
            Dim monthElement As IWebElement = Nothing
            Dim nbryears As Integer = CInt(CalendarRef.MonthIncDecValue / 12)
            Dim nbrmonths As Integer = 0
            If nbryears = 0 Then
                GoTo JumpMonths
            End If
            If CalendarRef.IsPositive(nbryears) Then
                For x As Integer = 1 To nbryears
                    yearObject = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.yearinc))
                    If Not yearObject Is Nothing Then
                        yearElement = DirectCast(yearObject, IWebElement)
                        yearElement.Click()
                        Threading.Thread.Sleep(50)
                    Else
                        Calendar_Found = False
                    End If
                Next
            Else
                For x As Integer = -1 To nbryears Step -1
                    yearObject = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.yeardec))
                    If Not yearObject Is Nothing Then
                        yearElement = DirectCast(yearObject, IWebElement)
                        yearElement.Click()
                        Threading.Thread.Sleep(50)
                    Else
                        Calendar_Found = False
                    End If
                Next
            End If
JumpMonths:
            nbrmonths = CalendarRef.MonthIncDecValue - (nbryears * 12)
            If nbrmonths > 0 Then
                For x As Integer = 1 To nbrmonths
                    'For x As Integer = 1 To CalendarRef.MonthIncDecValue
                    monthObject = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.monthinc))
                    If Not monthObject Is Nothing Then
                        monthElement = DirectCast(monthObject, IWebElement)
                        monthElement.Click()
                        Threading.Thread.Sleep(50)
                    Else
                        Calendar_Found = False
                    End If
                Next
            Else
                If nbrmonths < 0 Then
                    For x As Integer = -1 To nbrmonths Step -1
                        monthObject = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.monthdec))
                        If Not monthObject Is Nothing Then
                            monthElement = DirectCast(monthObject, IWebElement)
                            monthElement.Click()
                            Threading.Thread.Sleep(50)
                        Else
                            Calendar_Found = False
                        End If
                    Next
                End If
            End If

            'Check to see if the calendare is month and year have been set properly and report
            Dim checkDateCal As String = String.Empty
            Dim checkDatePassed As String = CalendarRef.TargetDate.ToString("MMMM") & ", " & CalendarRef.TargetYear
            Dim oCheckDate As New AutomationObject

            'oCheckDate.aObject = WaitForElement(driver, WaitForTime, By.XPath("//td[@id='pageTemplateForm:expectedDrawDateHeader']/table/tbody/tr/td[3]/div"))
            oCheckDate.aObject = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.CalHeader))
            If Not oCheckDate.aObject Is Nothing Then
                oCheckDate.aElement = DirectCast(oCheckDate.aObject, IWebElement)
                checkDateCal = oCheckDate.aElement.Text
            Else
                Calendar_Found = False
            End If



            If checkDatePassed = checkDateCal Then
                mosaicDll.Logger(senderMCPAction, "Actual Result| Month and Year Verified:" & checkDatePassed, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " Date value was not set properly."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Failed month and year check.  Expected " & checkDatePassed & " Calendar value was " & checkDateCal, myLogPath, , actionHistoryRecordID)
            End If

            'calculate the correct day index
            Dim tempMonth As Integer = CalendarRef.TargetMonth - 1
            If tempMonth = 0 Then
                'if target month is January, control will come inside.
                tempMonth = 12 'december
            End If

            Dim oZeroXPath As Object = WaitForElement(driver, WaitForTime, By.XPath(CalendarRef.dayZeroXPath))
            Dim oZeroPath_element As IWebElement = Nothing
            If Not oZeroXPath Is Nothing Then
                oZeroPath_element = DirectCast(oZeroXPath, IWebElement)
            End If

            Dim str As String = oZeroPath_element.Text
            Dim dayValueAtZeroIndex As Integer = Convert.ToInt32(str)


            If dayValueAtZeroIndex = 1 Then
                index = CalendarRef.TargetDay - 1
            Else
                Select Case tempMonth
                    Case 1, 3, 5, 7, 8, 10, _
                     12
                        If True Then
                            index = ((MAX_DAY_IN_A_MONTH - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                            Exit Select
                        End If
                    Case 2
                        If True Then
                            '
                            '	 * Separate case is needed for feb because Feb contains only 28(MAX_DAY_IN_A_MONTH - 3) days.
                            '	 
                            Dim bLeapYear As Boolean = Date.IsLeapYear(CalendarRef.TargetYear)

                            If bLeapYear Then
                                'index = (((MAX_DAY_IN_A_MONTH - 3) - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                                index = (((MAX_DAY_IN_A_MONTH - 2) - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                            Else
                                'index = (((MAX_DAY_IN_A_MONTH - 4) - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                                index = (((MAX_DAY_IN_A_MONTH - 3) - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                            End If

                            Exit Select
                        End If
                    Case Else
                        If True Then


                            index = (((MAX_DAY_IN_A_MONTH - 1) - dayValueAtZeroIndex) + CalendarRef.TargetDay)
                            Exit Select
                        End If
                End Select
            End If


            CalendarRef.DayIndex = index


            Dim DayObject As Object = Nothing
            Dim DayElement As IWebElement = Nothing
            Dim CompleteDayPath As String = CalendarRef.dayXPathPart1 & CalendarRef.DayIndex & CalendarRef.dayXPathPart2
            DayObject = WaitForElement(driver, WaitForTime, By.XPath(CompleteDayPath))
            If Not DayObject Is Nothing Then
                DayElement = DirectCast(DayObject, IWebElement)
                DayElement.Click()
            Else
                Calendar_Found = False
            End If

            If Calendar_Found Then
                mosaicDll.Logger(senderMCPAction, "Actual Result| Entered " & sDateParm.ToString & " for " & sCalendarName & " calendar field.", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " Could not enter " & sDateParm.ToString & " for " & sCalendarName & " calendar field."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter " & sDateParm.ToString & " for " & sCalendarName & " calendar field.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

        End If



    End Sub
    Public Function GetWordText(ByVal sStringValue As String) As String

        Dim objWord As New Word.Application
        Dim WordDoc As New Word.Document

        '0 = file not found
        '1 = File found
        '2 = error message

        Try
            If System.IO.File.Exists(sStringValue) Then


                WordDoc = objWord.Documents.Open(sStringValue.ToString())
                Dim sParagraph As New StringBuilder

                '  If you want the user to see it ... 
                objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objWord.Visible = False

                WordDoc.Select()
                objWord.Selection.Copy()
                GetWordText = "File found"


            Else
                GetWordText = "File not found"
            End If


        Catch ex As Exception
            GetWordText = ex.Message

        Finally
            WordDoc.Close(SaveChanges:=False)
            WordDoc = Nothing
            objWord.Quit()
            objWord = Nothing


        End Try



    End Function
    Public Function EnterTextArea(ByVal oTextArea As TextAreaParms, ByRef returnValues() As Object, ByVal actionHistoryRecordID As Integer, Optional ByVal sField As String = "") As Boolean
        ' Determines if data is a Word file - if so opens Word file and pastes contents in field; otherwise enters data provided in field
        Dim Did_it_work As Boolean = True

        System.Windows.Forms.Clipboard.Clear()
        Dim sRetVal As String = GetWordText(oTextArea.TextToEnter)
        If sRetVal.ToUpper = "FILE NOT FOUND" Then
            Dim sXpath As String = oTextArea.TextAreaXpath
            Dim oText As New AutomationObject
            oText.aObject = WaitForElement(driver, WaitForTime, By.XPath(sXpath))
            If Not oText.aObject Is Nothing Then
                oText.aElement = DirectCast(oText.aObject, IWebElement)
                oText.aElement.SendKeys(Keys.Control + "a")
                oText.aElement.SendKeys(Keys.Delete)
                oText.aElement.SendKeys(oTextArea.TextToEnter)
            End If
            ElseIf sRetVal.ToUpper = "FILE FOUND" Then
                Dim sXpath As String = oTextArea.TextAreaXpath
                Dim oText As New AutomationObject
                oText.aObject = WaitForElement(driver, WaitForTime, By.XPath(sXpath))
                If Not oText.aObject Is Nothing Then
                    oText.aElement = DirectCast(oText.aObject, IWebElement)
                    oText.aElement.SendKeys(Keys.Control + "a")
                    oText.aElement.SendKeys(Keys.Delete)

                    Dim myClipboardTest As String = System.Windows.Forms.Clipboard.GetText()
                    oText.aElement.SendKeys(Keys.Control + "v")

                    mosaicDll.Logger(senderMCPAction, "Actual Result|" & sField & " was entered from Word document.", myLogPath, , actionHistoryRecordID)
                End If
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " " & sRetVal
                mosaicDll.Logger(senderMCPAction, "Actual Result|Error from GetWordText Function:" & sRetVal, myLogPath, , actionHistoryRecordID)
                Did_it_work = False
            End If
            EnterTextArea = Did_it_work
    End Function
    Public Function EnterTextBox(ByVal oTextBox As TextBoxParms, ByRef returnValues() As Object, ByVal actionHistoryRecordID As Integer, Optional ByVal sField As String = "") As Boolean
        ' Enters text box using Xpath
        Dim Did_it_work As Boolean = True

        Dim sXpath As String = oTextBox.TextXpath
        Dim oText As New AutomationObject
        oText.aObject = WaitForElement(driver, WaitForTime, By.XPath(sXpath))
        If Not oText.aObject Is Nothing Then
            oText.aElement = DirectCast(oText.aObject, IWebElement)
            oText.aElement.SendKeys(Keys.Control + "a")
            oText.aElement.SendKeys(Keys.Delete)
            oText.aElement.SendKeys(oTextBox.TextStringToEnter)
            mosaicDll.Logger(senderMCPAction, "Actual Result| Entered " & sField & ": " & oTextBox.TextStringToEnter.ToString, myLogPath, , actionHistoryRecordID)
        Else
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " " & "Could not enter : & sfield"
            mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter: " & oTextBox.TextStringToEnter.ToString & "in " & sField, myLogPath, , actionHistoryRecordID)
            Did_it_work = False
        End If

        EnterTextBox = Did_it_work
    End Function
    Public Function CleanString(ByVal sCleanMe As String) As String


        sCleanMe = sCleanMe.Replace(" ", "")  'Remove Spaces
        sCleanMe = sCleanMe.ToUpper() 'Convert to upper case
        sCleanMe = sCleanMe.ToUpper() 'Remove leading and trailing spaces

        CleanString = sCleanMe

    End Function
#Region "Custom Selenium functions"
    Public Function ta_launch_browser(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            'open a new browser and open a blank page
            Select Case mosaicDll.dt(excelPath, iteration, "browser", "Browser").ToLower.Trim
                Case "firefox"
                    returnValues(1) = "Using Firefox browser."
                    driver = New Firefox.FirefoxDriver
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Firefox", myLogPath, , actionHistoryRecordID)
                Case "chrome"
                    returnValues(1) = "Using Chrome browser."
                    Dim ChromeOptn As New OpenQA.Selenium.Chrome.ChromeOptions
                    'Dim Service As OpenQA.Selenium.Remote.RemoteWebDriver()
                    ChromeOptn.BinaryLocation = (mosaicDll.dt(excelPath, iteration, "browser", "Path"))
                    'ChromeOptn.AddArguments("no-sandbox")
                    'ChromeOptn.AddArguments("--start-maximized")
                    ChromeOptn.AddArguments((mosaicDll.dt(excelPath, iteration, "browser", "Option")))
                    ' driver = New Chrome.ChromeDriver((mosaicDll.dt(excelPath, iteration, "browser", "DriverDirectoryPath")), ChromeOptn)
                    driver = New ChromeDriver((mosaicDll.dt(excelPath, iteration, "browser", "DriverDirectoryPath")), ChromeOptn)
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Chrome", myLogPath, , actionHistoryRecordID)
                Case Else
                    returnValues(1) = "Using Internet Explorer browser."
                    Dim options As New OpenQA.Selenium.IE.InternetExplorerOptions()
                    options.IntroduceInstabilityByIgnoringProtectedModeSettings = True
                    driver = New IE.InternetExplorerDriver((mosaicDll.dt(excelPath, iteration, "browser", "DriverDirectoryPath")), options)

                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Internet Explorer", myLogPath, , actionHistoryRecordID)
            End Select
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10))
            driver.Navigate().GoToUrl(mosaicDll.dt(excelPath, iteration, "url", "Target_page"))
            ' driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10))
            ' mosaicDll.Logger(senderMCPAction, "Actual Result|Navigated to " & mosaicDll.dt(excelPath, iteration, "url", "Target_page"), myLogPath, , actionHistoryRecordID)
            '  System.Threading.Thread.Sleep(500)
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_close_browser(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            'close the browser
            driver.Close()
            'driver.Quit()
            mosaicDll.Logger(senderMCPAction, "Actual Result|Closed Browser", myLogPath, , actionHistoryRecordID)
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_select_menu_option(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Dim LevelofMenu As String = mosaicDll.dt(excelPath, iteration, "menu_selection", "LevelofMenu").Trim
        Try
            'locate an element and click it
            Dim oMenu As New AutomationObject
            Dim Xpath As String = mosaicDll.dt(excelPath, iteration, "menu_selection", "Link_Path").Trim

            oMenu.aObject = WaitForElement(driver, WaitForTime, By.XPath(Xpath))
            If Not oMenu.aObject Is Nothing Then
                oMenu.aElement = DirectCast(oMenu.aObject, IWebElement)
                oMenu.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked " & mosaicDll.dt(excelPath, iteration, "menu_selection", "Tab_name"), myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " Unable to Click Menu Tab"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Unable to Click " & mosaicDll.dt(excelPath, iteration, "menu_selection", "Tab_name"), myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If
            Return returnValues

        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_request_form_link(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Dim bClickedlink As Boolean = False
        Try
            'click a hyperlink
            'driver.FindElement(By.LinkText(mosaicDll.dt(excelPath, iteration, "menu_selection", "Link_Path"))).Click()
            Dim Link As String = mosaicDll.dt(excelPath, iteration, "menu_selection", "Link_Path")
            Dim olink As New AutomationObject
            olink.aObject = WaitForElement(driver, WaitForTime, By.LinkText(Link))
            If Not olink.aObject Is Nothing Then
                olink.aElement = DirectCast(olink.aObject, IWebElement)
                olink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked " & Link.ToString, myLogPath, , actionHistoryRecordID)
                bClickedlink = True

            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click link text"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click " & Link.ToString, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
                Return returnValues
            End If
            Return returnValues
        Catch ex As Exception
            If bClickedlink = False Then
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
                mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
                mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
                Return returnValues
            End If
            Return returnValues
        End Try
    End Function
    Public Function ta_enter_data(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)

        Dim sfirstname As String = mosaicDll.dt(excelPath, iteration, "form_data", "First_Name")
        Dim slastname As String = mosaicDll.dt(excelPath, iteration, "form_data", "Last_Name")
        Dim scompany As String = mosaicDll.dt(excelPath, iteration, "form_data", "Company")
        Dim sTitle As String = mosaicDll.dt(excelPath, iteration, "form_data", "Title")
        Dim saddress1 As String = mosaicDll.dt(excelPath, iteration, "form_data", "Address1")
        Dim saddress2 As String = mosaicDll.dt(excelPath, iteration, "form_data", "Address2")
        Dim scity As String = mosaicDll.dt(excelPath, iteration, "form_data", "City")
        Dim sstate As String = mosaicDll.dt(excelPath, iteration, "form_data", "State")
        Dim szipcode As String = mosaicDll.dt(excelPath, iteration, "form_data", "Zip_Code")
        Dim scountry As String = mosaicDll.dt(excelPath, iteration, "form_data", "Country")
        Dim sphone As String = mosaicDll.dt(excelPath, iteration, "form_data", "Phone")
        Dim semail As String = mosaicDll.dt(excelPath, iteration, "form_data", "Email")
        Dim iFailcount As Integer = 0
        Dim bresult As Boolean
        Dim TextBox As New TextBoxParms
        Try

            'enter first name

            TextBox.TextStringToEnter = sfirstname
            TextBox.TextXpath = "//input[@id='FormsEditField1']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "First Name")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter last name

            TextBox.TextStringToEnter = slastname
            TextBox.TextXpath = "//input[@id='FormsEditField9']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Last Name")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter company

            TextBox.TextStringToEnter = scompany
            TextBox.TextXpath = "//input[@id='FormsEditField2']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Company")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter title

            TextBox.TextStringToEnter = sTitle
            TextBox.TextXpath = "//input[@id='FormsEditField10']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Title")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter address 1

            TextBox.TextStringToEnter = saddress1
            TextBox.TextXpath = "//input[@id='FormsEditField4']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Address Line One")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If
            'enter address 2

            TextBox.TextStringToEnter = saddress2
            TextBox.TextXpath = "//input[@id='FormsEditField5']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Address Line Two")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter city

            TextBox.TextStringToEnter = scity
            TextBox.TextXpath = "//input[@id='FormsEditField8']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Address Line Two")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter state

            TextBox.TextStringToEnter = sstate
            TextBox.TextXpath = "//input[@id='FormsEditField11']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Address Line Two")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter zip code

            TextBox.TextStringToEnter = szipcode
            TextBox.TextXpath = "//input[@id='FormsEditField12']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Zip Code")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter country

            TextBox.TextStringToEnter = scountry
            TextBox.TextXpath = "//input[@id='FormsEditField6']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Country")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter phone

            TextBox.TextStringToEnter = sphone
            TextBox.TextXpath = "//input[@id='FormsEditField7']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Phone")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            'enter email

            TextBox.TextStringToEnter = semail
            TextBox.TextXpath = "//input[@id='FormsEditField3']"

            bresult = EnterTextBox(TextBox, returnValues, actionHistoryRecordID, "Phone")
            If bresult = False Then
                iFailcount = iFailcount + 1
            End If

            Return returnValues

        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_submit(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Dim sxpath As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Selenium_Path")
        Dim sbutton As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Action_name")
        Try
            'click the Submit button
            Dim oSubmit As New AutomationObject
            oSubmit.aObject = WaitForElement(driver, WaitForTime, By.XPath(sxpath))
            If Not oSubmit.aObject Is Nothing Then
                oSubmit.aElement = DirectCast(oSubmit.aObject, IWebElement)
                oSubmit.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked " & sbutton, myLogPath, , actionHistoryRecordID)
                'System.Threading.Thread.Sleep(500)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click " & sbutton
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click " & sbutton, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
                Return returnValues
            End If

            'driver.FindElement(By.Id(mosaicDll.dt(excelPath, iteration, "action_buttons", "Selenium_Path"))).Submit()
            'mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked " & mosaicDll.dt(excelPath, iteration, "action_buttons", "Action_name"), myLogPath, , actionHistoryRecordID)
            ' System.Threading.Thread.Sleep(1500)
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_capture_screenshot(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Dim snagItDirectory As String = ""
        Dim snagItFileName As String = "" 'file name w/o the extension
        Dim snagItFileExtension As String = ""  'extension of the image file
        Try
            'get the directory for the SnagIt path
            snagItDirectory = My.Computer.FileSystem.CombinePath(My.Computer.FileSystem.GetParentPath(excelPath), "Screenshots")

            'get the current date/time and use that as the snagit file name
            snagItFileName = My.Computer.Clock.LocalTime.Year.ToString & _
                My.Computer.Clock.LocalTime.Month.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Day.ToString.PadLeft(2, "0"c) & _
                "_" & _
                My.Computer.Clock.LocalTime.Hour.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Minute.ToString.PadLeft(2, "0"c) & _
                My.Computer.Clock.LocalTime.Second.ToString.PadLeft(2, "0"c)

            'set snagit to take a desktop screenshot and store it as a file
            snagIt.Input = SNAGITLib.snagImageInput.siiDesktop
            snagIt.Output = SNAGITLib.snagImageOutput.sioFile

            'set the screenshot properties and save location
            snagIt.OutputImageFile.FileNamingMethod = SNAGITLib.snagOuputFileNamingMethod.sofnmFixed
            snagIt.OutputImageFile.Directory = snagItDirectory
            snagIt.OutputImageFile.Filename = snagItFileName
            snagIt.OutputImageFile.FileType = SNAGITLib.snagImageFileType.siftPNG
            Select Case snagIt.OutputImageFile.FileType
                Case snagImageFileType.siftPNG
                    snagItFileExtension = ".png"
                Case snagImageFileType.siftJPEG
                    snagItFileExtension = ".jpeg"
                Case snagImageFileType.siftGIF
                    snagItFileExtension = ".gif"
                Case snagImageFileType.siftBMP
                    snagItFileExtension = ".bmp"
                Case Else
                    snagItFileExtension = ""
            End Select

            'don't show snagit preview window and cursor
            snagIt.EnablePreviewWindow = False
            snagIt.IncludeCursor = False

            'take a snapshot, check for snagit license
            Try
                snagIt.Capture()
                mosaicDll.RSTAR_AddScreenshot(CLng(myExecutableHistoryRecordID), 5, CLng(actionHistoryRecordID), My.Computer.FileSystem.CombinePath(snagItDirectory, snagItFileName & snagItFileExtension), "") 'Use 5 to attach a screenshot to an Action in the history table.
                mosaicDll.Logger(senderMCPAction, "Screenshot|" & My.Computer.FileSystem.CombinePath(snagItDirectory, snagItFileName & snagItFileExtension), myLogPath, , actionHistoryRecordID)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Screenshot captured.", myLogPath, , actionHistoryRecordID)
            Catch ex As Exception
                returnValues(0) = executionStatusFailed
                If snagIt.LastError = SNAGITLib.snagError.serrSnagItExpired Then
                    returnValues(1) = "Unable to capture a screenshot: SnagIt evaluation has expired."
                Else
                    returnValues(1) = "Unable to capture a screenshot: " & ex.Message
                End If
                Return returnValues
            End Try
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_verification_action(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            Dim expectedText As String = mosaicDll.dt(excelPath, iteration, "form_data", "Verification_text")
            Dim sxpath As String = mosaicDll.dt(excelPath, iteration, "form_data", "Verfication_XPATH")
            Dim oCheckforobject As New AutomationObject
            If Not String.IsNullOrEmpty(sxpath) Then
                oCheckforobject.aObject = WaitForElement(driver, WaitForTime, By.XPath(sxpath))
                If Not oCheckforobject.aObject Is Nothing Then
                    oCheckforobject.aElement = DirectCast(oCheckforobject.aObject, IWebElement)
                    Dim actualtext As String = oCheckforobject.aElement.Text
                    If CleanString(actualtext) = CleanString(expectedText) Then
                        mosaicDll.Logger(senderMCPAction, "Actual Result|Expected Text Found: " & expectedText & ".", myLogPath, , actionHistoryRecordID)
                    Else
                        returnValues(1) = "Element found but wrong Text: " & actualtext.ToString & "'."
                        mosaicDll.Logger(senderMCPAction, "Actual Result|Expected Text: " & expectedText & " found text" & actualtext & ".", myLogPath, , actionHistoryRecordID)
                    End If
                Else
                    returnValues(1) = "Specified Element not found."
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Could not find element: " & expectedText & ".", myLogPath, , actionHistoryRecordID)
                End If
            Else
                oCheckforobject.aObject = WaitForElement(driver, WaitForTime, By.LinkText(expectedText))
                If Not oCheckforobject.aObject Is Nothing Then
                    returnValues(1) = "Expected Link Text Found: " & expectedText & "'."
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Expected Link Text Found: " & expectedText & ".", myLogPath, , actionHistoryRecordID)
                Else
                    returnValues(0) = executionStatusFailed
                    returnValues(1) = "Expected link was not found."
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Expected Link: " & expectedText & " was not found.", myLogPath, , actionHistoryRecordID)
                    CaptureScreenshot(actionHistoryRecordID)
                End If
            End If
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_drop_down_list(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            'create an object for the dropdown, then select the text
            Dim item As SelectElement = New SelectElement(driver.FindElement(By.Id(mosaicDll.dt(excelPath, iteration, "selection_down", "list_Name"))))
            item.SelectByText(mosaicDll.dt(excelPath, iteration, "selection_down", "Selection_value"))

            mosaicDll.Logger(senderMCPAction, "Actual Result|Selected from dropdown: " & mosaicDll.dt(excelPath, iteration, "selection_down", "Selection_value"), myLogPath, , actionHistoryRecordID)
            System.Threading.Thread.Sleep(2000) 'useful if filling out multiple lists that are dependent on one another
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_select_radio_button(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            Dim oClickRadial As New AutomationObject
            oClickRadial.aObject = WaitForElement(driver, WaitForTime, By.XPath("//input[@id='nof']"))
            If Not oClickRadial.aObject Is Nothing Then
                oClickRadial.aElement = DirectCast(oClickRadial.aObject, IWebElement)
                oClickRadial.aElement.Click()
                returnValues(1) = CStr(returnValues(1)) & "Clicked Radial Button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked Radial Button", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not Click Radial Button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not Click Radial Button", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_set_checkbox(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            Dim oSetCheckBox As New AutomationObject
            oSetCheckBox.aObject = WaitForElement(driver, WaitForTime, By.Id("paf"))
            If Not oSetCheckBox.aObject Is Nothing Then
                oSetCheckBox.aElement = DirectCast(oSetCheckBox.aObject, IWebElement)
                If oSetCheckBox.aElement.Selected = True Then
                    'proper state
                Else
                    oSetCheckBox.aElement.Click()
                End If
                returnValues(1) = CStr(returnValues(1)) & "Clicked Check Box"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked Check Box", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not Click Check Box"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not Click Check Box", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_enter_zip_code(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try
            'enter zip code
            Dim zipCode As String = mosaicDll.dt(excelPath, iteration, "form_data", "Zip_Code")
            Dim oZipCode As New AutomationObject
            oZipCode.aObject = WaitForElement(driver, WaitForTime, By.Id("zcnew"))
            If Not oZipCode.aObject Is Nothing Then
                oZipCode.aElement = DirectCast(oZipCode.aObject, IWebElement)
                oZipCode.aElement.Clear()
                oZipCode.aElement.SendKeys(zipCode)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered Zip Code: " & zipCode, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter Zip Code"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter Zip Code: " & zipCode, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            mosaicDll.Logger(senderMCPAction, "Actual Result|Entered Zip Code: " & zipCode, myLogPath, , actionHistoryRecordID)
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_click_button(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Dim sbuttonxpath As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Selenium_Path")
        Dim sactionname As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Action_name")
        Try
            'click the button
            Dim oButton As New AutomationObject
            oButton.aObject = WaitForElement(driver, WaitForTime, By.Id("zcnew"))
            If Not oButton.aObject Is Nothing Then
                oButton.aElement = DirectCast(oButton.aObject, IWebElement)
                oButton.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked " & sactionname, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click: " & sactionname, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_enter_search(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)

        Try
            Dim SearchXpath As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Selenium_Path")
            Dim Searchvalue As String = mosaicDll.dt(excelPath, iteration, "search_criteria", "Search_Criteria")
            Dim sActionname As String = mosaicDll.dt(excelPath, iteration, "action_buttons", "Action_name")
            Dim oSearch As New AutomationObject
            'enter search value
            oSearch.aObject = WaitForElement(driver, WaitForTime, By.XPath(SearchXpath))
            If Not oSearch.aObject Is Nothing Then
                oSearch.aElement = DirectCast(oSearch.aObject, IWebElement)
                oSearch.aElement.SendKeys(Searchvalue)
                oSearch.aElement.SendKeys(Keys.Enter)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered Search Criteria: " & Searchvalue, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & " Could not enter search value"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not enter search value: " & Searchvalue, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'driver.FindElement(By.XPath(mosaicDll.dt(excelPath, iteration, "action_buttons", "Selenium_Path"))).Submit()
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function

    '1 - SSO Login
    Public Function ta_NAV_navigate_to_ahaonlinestore(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)

        'Dim strWhichBrowser As String = mosaicDll.dt(excelPath, iteration, "browser", "Browser").ToLower.Trim
        Dim strWhichBrowser As String = "Internet Explorer"

        Try
            'open a new browser and open a blank page
            Select Case strWhichBrowser
                Case "firefox"
                    returnValues(1) = "Using Firefox browser."
                    driver = New Firefox.FirefoxDriver
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Firefox", myLogPath, , actionHistoryRecordID)
                Case "chrome"
                    returnValues(1) = "Using Chrome browser."
                    Dim ChromeOptn As New OpenQA.Selenium.Chrome.ChromeOptions
                    'Dim Service As OpenQA.Selenium.Remote.RemoteWebDriver()
                    ChromeOptn.BinaryLocation = (mosaicDll.dt(excelPath, iteration, "browser", "Path"))
                    'ChromeOptn.AddArguments("no-sandbox")
                    'ChromeOptn.AddArguments("--start-maximized")
                    ChromeOptn.AddArguments((mosaicDll.dt(excelPath, iteration, "browser", "Option")))
                    ' driver = New Chrome.ChromeDriver((mosaicDll.dt(excelPath, iteration, "browser", "DriverDirectoryPath")), ChromeOptn)
                    driver = New ChromeDriver((mosaicDll.dt(excelPath, iteration, "browser", "DriverDirectoryPath")), ChromeOptn)
                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Chrome", myLogPath, , actionHistoryRecordID)
                Case Else
                    returnValues(1) = "Using Internet Explorer browser."
                    Dim options As New OpenQA.Selenium.IE.InternetExplorerOptions()
                    options.IntroduceInstabilityByIgnoringProtectedModeSettings = True
                    ' driver = New IE.InternetExplorerDriver((mosaicDll.dt(excelPath, iteration, "logo_&_branding_info", "url")), options)
                    driver = New IE.InternetExplorerDriver(("C:\Projects\SeleniumMCP\Selenium_Drivers"), options)


                    mosaicDll.Logger(senderMCPAction, "Actual Result|Launched Internet Explorer", myLogPath, , actionHistoryRecordID)
            End Select

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10))
            driver.Navigate().GoToUrl(mosaicDll.dt(excelPath, iteration, "logo_&_branding_info", "url"))
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10))
 
            mosaicDll.Logger(senderMCPAction, "Actual Result|Navigated to " & mosaicDll.dt(excelPath, iteration, "logo_&_branding_info", "url"), myLogPath, , actionHistoryRecordID)
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            CaptureScreenshot(actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_click_eweb_login_link(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to login button
            Dim strLoginXPath As String = "//div[@id='LoginDetail']/div/a[1]"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strLoginXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_enter_eweb_login_info(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Setup XPAths
            Dim sEmailXPath As String = "//input[@id='userid']"
            Dim sPassXPath As String = "//input[@id='password']"
            Dim sRememberMeXPath As String = "//input[@id='remember_me']"


            'Get Data Values
            Dim sEmail As String = mosaicDll.dt(excelPath, iteration, "individual_info", "primary_email")
            Dim sPassword As String = mosaicDll.dt(excelPath, iteration, "login_info", "password")
            Dim sRememeberMe As String = mosaicDll.dt(excelPath, iteration, "login_info", "remember_me_check_box").ToLower.Trim


            'Enter the email address
            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sEmailXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.SendKeys(Keys.Control + "a")
                oLink.aElement.SendKeys(sEmail)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the email address. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the email address."
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter the email address: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sPassXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.SendKeys(Keys.Control + "a")
                oLink.aElement.SendKeys(sPassword)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the password. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the password"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter the email address: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sRememberMeXPath))
            'If Not oLink.aObject Is Nothing Then
            '    oLink.aElement = DirectCast(oLink.aObject, IWebElement)
            '    oLink.aElement.Click()
            '    mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the email address. ", myLogPath, , actionHistoryRecordID)
            'Else
            '    returnValues(0) = executionStatusFailed
            '    returnValues(1) = CStr(returnValues(1)) & "Could not enter the email address."
            '    mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter the email address: ", myLogPath, , actionHistoryRecordID)
            '    CaptureScreenshot(actionHistoryRecordID)
            'End If

            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_submit_valid_eweb_login(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            Dim sLoginButtonXPath As String = "//input[@id='login']"

            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sLoginButtonXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login button", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login button: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_click_eweb_logout_link(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            Dim sLogouButtonXPath As String = "Fill in Value Here"

            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sLogouButtonXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login button", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login button: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If
            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function

    '2 - New User Registration
    Public Function ta_NAV_click_eweb_register_link(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to login button
            Dim strRegisterXPath As String = "//div[@id='LoginDetail']/div/a[2]"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strRegisterXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the register link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the register link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_enter_email_address(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Setup XPAths
            Dim sEmailXPath As String = "/html/body/form/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr/td/div/span[2]/div/table/tbody/tr[1]/td[2]/input"


            'Get Data Values
            Dim sEmail As String = mosaicDll.dt(excelPath, iteration, "individual_info", "primary_email")


            'Enter the email address
            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sEmailXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.SendKeys(Keys.Control + "a")
                oLink.aElement.SendKeys(sEmail)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the email address. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the email address"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not enter the email address: ", myLogPath, , actionHistoryRecordID)
                ' CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_REG_submit_valid_new_account_email(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to login button
            Dim sSubmitButtonID As String = "SearchBtn"
            Dim sSubmitButtonXPath As String = "//input[@id='SearchBtn']"
            Dim sMessageXPath As String = "//span[@id='resultheader']/p"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sSubmitButtonXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the registration link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the registration link."
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the register link: ", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If

            'Message to Check ask Florence about same should be templated
            Dim sMessageText As String = mosaicDll.dt(excelPath, iteration, "individual_info", "AR_Verification_Message")

            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sMessageXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                If oLink.aElement.Text.ToLower.Trim() = sMessageText.ToLower.Trim() Then
                    mosaicDll.Logger(senderMCPAction, "Actual Result|The expected message was displayed.", myLogPath, , actionHistoryRecordID)
                Else
                    mosaicDll.Logger(senderMCPAction, "Actual Result|The expected message was  not displayed.", myLogPath, , actionHistoryRecordID)
                End If

            Else
                returnValues(0) = executionStatusFailedAbort
                returnValues(1) = CStr(returnValues(1)) & "Could not locate the text for comparison of the expected message text."
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the register link: ", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If

            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailedAbort
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_REG_click_register_link_on_new_acct_reg_screen(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to login button
            Dim strRegisterParagraphLink As String = "//span[@id='resultheader']/p/a/b"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strRegisterParagraphLink))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the register link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the register link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_REG_enter_new_acct_reg_info(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Setup XPAths
            Dim sEmailXPath As String = "//input[@id='userid']"
            Dim sPassXPath As String = "//input[@id='password']"
            Dim sRememberMeXPath As String = "//input[@id='remember_me']"
            Dim sPrefixXPath As String = "//select[@id='ind_prf_code']"
            Dim sSuffixXPath As String = "//select[@id='ind_sfx_code']"
            'Get Data Values
            Dim sPrefix As String = mosaicDll.dt(excelPath, iteration, "individual_info", "prefix")
            Dim sFirstName As String = mosaicDll.dt(excelPath, iteration, "individual_info", "first_name")
            Dim sMiddleName As String = mosaicDll.dt(excelPath, iteration, "individual_info", "middle_name")
            Dim sLastName As String = mosaicDll.dt(excelPath, iteration, "individual_info", "last_name")
            Dim sSuffix As String = mosaicDll.dt(excelPath, iteration, "individual_info", "suffix")

            Dim saddress_line_1 As String = mosaicDll.dt(excelPath, iteration, "address_info", "address_line_1")
            Dim saddress_line_2 As String = mosaicDll.dt(excelPath, iteration, "address_info", "address_line_2")
            Dim sCity As String = mosaicDll.dt(excelPath, iteration, "address_info", "city")
            Dim sState As String = mosaicDll.dt(excelPath, iteration, "address_info", "state")
            Dim sZip As String = mosaicDll.dt(excelPath, iteration, "address_info", "zip")
            Dim sCountry As String = mosaicDll.dt(excelPath, iteration, "address_info", "country")
            Dim sEmailAddress As String = mosaicDll.dt(excelPath, iteration, "individual_info", "primary_email")
            Dim sCurrentPassword As String = mosaicDll.dt(excelPath, iteration, "login_info", "password")


            'Select prefix 
            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sPrefixXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                Dim selectElement As SelectElement = New SelectElement(oLink.aElement)
                selectElement.SelectByValue(sPrefix)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Set the combo box to: " & sPrefix, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not set the combo box value"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not set the combo box value to: " & sPrefix, myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Enter first Nname
            Dim oField As New AutomationObject
            Dim sFirstNameID As String = "ind_first_name"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sFirstNameID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sFirstName)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered first name: ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "could not enter the first name."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the first name.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            'Enter middle name
            Dim sMiddleNameID As String = "ind_mid_name"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sMiddleNameID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sMiddleName)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered middle name: ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "could not enter the middle name."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the middle name.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Enter last name
            Dim sLastNameID As String = "ind_last_name"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sLastNameID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sLastName)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered last name: ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the last name."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the last name.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Select suffix(combo box)
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(sSuffixXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                Dim selectElement As SelectElement = New SelectElement(oLink.aElement)
                selectElement.SelectByValue(sSuffix)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Set the combo box to: " & sSuffix, myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not set the combo box value"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not set the combo box value to: " & sSuffix, myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If


            'Enter address line 1 and 2
            Dim sAddress1ID As String = "adr_line1"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sAddress1ID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(saddress_line_1)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered address line 1.", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the address line 1."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the address line 1.", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If

            Dim sAddress2ID As String = "adr_line2"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sAddress2ID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(saddress_line_2)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered address line 2", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the address line 2"
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the address line 2", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Enter city state zip
            Dim sCityID As String = "adr_city"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sCityID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sCity)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered city. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "could not enter the city."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the city", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Enter city state zip
            Dim sStateID As String = "adr_state"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sStateID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sState)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered state. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the state."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the state.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            'Enter city state zip
            Dim sZipID As String = "adr_post_code"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sZipID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sZip)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the postal code: ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the postal code."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the postal code.", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            Dim sEmailID As String = "eml_address"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sEmailID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sEmailAddress)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Email address has been entered. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the email address."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the email address.", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If

            Dim sPassword1ID As String = "AHA_Password1"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sPassword1ID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sCurrentPassword)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the initial password. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the initial password."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the initial password.", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If

            Dim sPassword2ID As String = "AHA_Password2"
            oField.aObject = WaitForElement(driver, WaitForTime, By.Id(sPassword2ID))
            If Not oField.aObject Is Nothing Then
                oField.aElement = DirectCast(oField.aObject, IWebElement)
                oField.aElement.SendKeys(Keys.Control + "a")
                oField.aElement.SendKeys(sCurrentPassword)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Entered the password verification ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not enter the password verification."
                mosaicDll.Logger(senderMCPAction, "Actual Result| Could not enter the password verification.", myLogPath, , actionHistoryRecordID)
                'CaptureScreenshot(actionHistoryRecordID)
            End If



            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_REG_submit_valid_new_acct_reg_info_wo_linking_to_org(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id("SubmitBtn"))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the request new account button. ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the request new account button."
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the request new account button. ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            'oLink.aObject = WaitForElement(driver, WaitForTime, By.Id("SubmitBtn"))
            'If Not oLink.aObject Is Nothing Then
            '    oLink.aElement = DirectCast(oLink.aObject, IWebElement)
            '    oLink.aElement.Click()
            '    mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the request new account button. ", myLogPath, , actionHistoryRecordID)
            'Else
            '    returnValues(0) = executionStatusFailed
            '    returnValues(1) = CStr(returnValues(1)) & "Could not click the request new account button."
            '    mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the request new account button. ", myLogPath, , actionHistoryRecordID)
            '    CaptureScreenshot(actionHistoryRecordID)
            'End If

            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_STOR_click_join_link_logged_in_no_questions(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to Login
            Dim strLoginXPath As String = "//div[@id='LoginDetail']/div/a[1]"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strLoginXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function

    '3
    Public Function ta_STOR_submit_valid_membership_info(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to Continue Button
            Dim strContinueXPath As String = "//div[@id='LoginDetail']/div/a[1]"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strContinueXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_STOR_click_add_to_cart_button_on_membership_screen(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to Add to Cart button
            Dim strCartXPath As String = "//div[@id='LoginDetail']/div/a[1]"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.XPath(strCartXPath))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could Not click login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function

    '4 Memnebership Renewal
    Public Function ta_SSO_click_eweb_login_link(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to renewal button
            Dim strRenewXPathID As String = "login"


            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id(strRenewXPathID))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the login button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_enter_eweb_login_info(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to renewal button
            Dim sEmailID As String = "Email"
            Dim sPasswordID As String = "Password"
            Dim sRememberMeID As String = "remember_me"

            'Get the Data
            Dim sEmail As String = mosaicDll.dt(excelPath, iteration, "individual_info", "primary_email")
            Dim sPassword As String = mosaicDll.dt(excelPath, iteration, "individual_info", "password")
            Dim sRememberMe As String = mosaicDll.dt(excelPath, iteration, "individual_info", "remember_me_check_box")



            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id(sEmailID)
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.SendKeys(Keys.Control + "a")
                oLink.aElement.SendKeys(sEmail)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the login button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id(sEmailID)
            If Not oLink.aObject Is Nothing Then
                oLink.aElement.SendKeys(Keys.Control + "a")
                oLink.aElement.SendKeys(sPassword)
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the login button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If

            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id(sEmailID)
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                'coompare check box values
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the login button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function
    Public Function ta_SSO_submit_valid_eweb_login(ByVal mcpParameters() As Object) As Object()
        Dim returnValues() As Object = New Object() {executionStatusPassed, ""}
        Dim excelPath As String = CType(mcpParameters(0), String)
        Dim iteration As Integer = CType(mcpParameters(1), Integer)
        Dim actionHistoryRecordID As Integer = CType(mcpParameters(2), Integer)
        Try

            'Path to renewal button
            Dim sLoginID As String = "login"

            Dim oLink As New AutomationObject
            oLink.aObject = WaitForElement(driver, WaitForTime, By.Id(sLoginID))
            If Not oLink.aObject Is Nothing Then
                oLink.aElement = DirectCast(oLink.aObject, IWebElement)
                oLink.aElement.Click()
                mosaicDll.Logger(senderMCPAction, "Actual Result|Clicked the login link ", myLogPath, , actionHistoryRecordID)
            Else
                returnValues(0) = executionStatusFailed
                returnValues(1) = CStr(returnValues(1)) & "Could not click the login button"
                mosaicDll.Logger(senderMCPAction, "Actual Result|Could not click the login link: ", myLogPath, , actionHistoryRecordID)
                CaptureScreenshot(actionHistoryRecordID)
            End If


            Return returnValues
        Catch ex As Exception
            returnValues(0) = executionStatusFailed
            returnValues(1) = CStr(returnValues(1)) & " Exception: " & ex.Message
            mosaicDll.Logger(senderMCPAction, exceptionErrorMsg, myLogPath, , actionHistoryRecordID)
            mosaicDll.Logger(senderMCPAction, "Error|Exception: '" & ex.Message & "'", myLogPath, , actionHistoryRecordID)
            Return returnValues
        End Try
    End Function


#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
