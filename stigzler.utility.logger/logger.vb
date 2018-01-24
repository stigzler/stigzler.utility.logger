Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.CompilerServices

Public Class Log

#Region "Properties"

    Private _filename As String = Application.StartupPath & "\" & AppName() & ".log"
    Private _sysdets As Boolean = False
    Private _appname As String = AppName()
    Private _appvers As String = "Unset"
    Private _logtype As LogType = LogType.OverwriteOld
    Private _loglevel As Integer = 1
    Private _datetimeformat As String = "H:mm:ss-ff"
    Private _useconsole As Boolean = False

    ''' <summary>
    ''' Full filename of log file including path, name and extention
    ''' </summary>
    Public ReadOnly Property Filename As String
        Get
            Return _filename
        End Get
    End Property

    ''' <summary>
    ''' Whether to include client system details at the start of the log
    ''' </summary>
    Public Property IncludeSystemDetails As Boolean
        Get
            Return _sysdets
        End Get
        Set(value As Boolean)
            _sysdets = value
        End Set
    End Property

    ''' <summary>
    ''' What type of log to keep.
    ''' </summary>
    Public Property LogType As LogType
        Get
            Return _logtype
        End Get
        Set(value As LogType)
            _logtype = value
        End Set
    End Property

    ''' <summary>
    ''' Name of the application
    ''' </summary>
    Public Property ApplicationName As String
        Get
            Return _appname
        End Get
        Set(value As String)
            _appname = value
        End Set
    End Property

    ''' <summary>
    ''' Application Version
    ''' </summary>
    Public Property ApplicationVersion As String
        Get
            Return _appvers
        End Get
        Set(value As String)
            _appvers = value
        End Set
    End Property

    ''' <summary>
    ''' Value or 1 to 5. What level of logging lines to allow up to. For example, "3" will allow logging lines of levels 1, 2 and 3 to be written. Logging line of 4 and 5 will not be written.
    ''' </summary>
    ''' <remarks>Typically, this is used to allow your user to set different levels of logging to accommodate things such as debugging and verbose logging.</remarks>
    ''' <value>1</value>
    Public Property LogLevel As Integer
        Get
            Return _loglevel
        End Get
        Set(value As Integer)
            If value < 1 Then
                _loglevel = 1
            ElseIf value > 5 Then
                _loglevel = 5
            Else
                _loglevel = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Date/time format string to use for the line prefix.
    ''' </summary>
    Public Property DateTimeFormat As String
        Get
            Return _datetimeformat
        End Get
        Set(value As String)
            _datetimeformat = value
        End Set
    End Property



#End Region

    ''' <summary>
    ''' Instantiates a new log. Default save as [application name].log in application directory
    ''' </summary>
    Sub New()

    End Sub

    ''' <summary>
    ''' Start logging using parameters set.
    ''' </summary>
    Public Sub StartLogging()

        If _logtype = LogType.OverwriteOld Then
            If File.Exists(_filename) Then
                File.Delete(_filename)

            End If
        Else
            Write_Line("===============================================================")
        End If

        Write_Line([String].Format("LOG: {0} (Version: {1})", _appname, _appvers))
        Write_Line("Log Started: " & DateTime.Now.ToString("dd MMMM, yyyy. H:mm:ss-ff"))
        Write_Line("Log Level: " & _loglevel & " (1 is basic, 5 is verbose).")
        Write_Line("Log Type: " & _logtype.ToString)

        If _sysdets Then
            Get_System_Info()
        End If

    End Sub

    ''' <summary>
    ''' Clears any logging in the existing file
    ''' </summary>
    Public Sub ClearLog()
        If File.Exists(_filename) Then
            File.Delete(_filename)
        End If
    End Sub

    Private Sub Dbug(line As String)
        Debug.WriteLine(line)
    End Sub

    ''' <summary>
    ''' Returns an automatically obtained applicaiton name.
    ''' </summary>
    Private Function AppName() As String
        Dim location As String = Environment.GetCommandLineArgs()(0)
        AppName = Path.GetFileNameWithoutExtension(location)
    End Function

    ''' <summary>
    ''' Write lines to your log file.
    ''' </summary>
    ''' <param name="Log_Entry">The entry to write to the log file.</param>
    Private Sub Write_Line(ByVal Log_Entry As String)
        Dbug("Writing line")



        'Write the line to the log file:
        Try
            Using sw As StreamWriter = System.IO.File.AppendText(_filename)
                sw.WriteLine([String].Format("{0} | {1}", DateTime.Now.ToString(_datetimeformat), Log_Entry))
                sw.Flush()
            End Using


        Catch ex As Exception
            'MessageBox.Show(ex.Message, "WRITE ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub



    ''' <summary>
    ''' Set the logfile name and location. Use the full filepath including the extention. Defaults to known valid or Appname.log in application folder on directory not existing.
    ''' </summary>
    Public Sub SetLogFile(ProposedFile As String)
        If Directory.Exists(Path.GetDirectoryName(ProposedFile)) Then
            _filename = ProposedFile
        End If
    End Sub

    ''' <summary>
    ''' Write a line in the log
    ''' </summary>
    ''' <param name="text">The text to print to a log line</param>
    Public Sub Log(text As String)

        Write_Line(text)

    End Sub

    ''' <summary>
    ''' Write a line in the log at a certain log level
    ''' </summary>
    ''' <param name="Text">The text to print to a log line</param>
    ''' <param name="loglevel">The log level of the line</param>
    Public Sub Log(Text As String, loglevel As Integer)

        If loglevel > _loglevel Then Return 'returns if this logline at higher level of detail than global setting

        Write_Line(Text)

    End Sub

    ''' <summary>
    ''' Write a line in the log
    ''' </summary>
    ''' <param name="text">The text to print to a log line</param>
    ''' <param name="info">Additional info to include in the log</param>
    Public Sub Log(text As String, info As Info,
                   <CallerMemberName> Optional memberName As String = "",
                   <CallerFilePath> Optional filename As String = "")
        Dbug("additonal info sub")
        Dim outline As String = Nothing

        Select Case info
            Case Info.CallingClass
                outline = Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & text
            Case Info.CallingMethod
                outline = memberName.ToUpper & " | " & text
            Case Info.CallingClassAndMethod
                outline = Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & memberName.ToUpper & " | " & text
        End Select

        Write_Line(outline)

    End Sub

    ''' <summary>
    '''  Write a line in the log
    '''  </summary>
    '''  <param name="text">The text to print to a log line</param>
    '''  <param name="info">Additional info to include in the log</param>
    ''' <param name="loglevel">Write a line in the log at a certain log level</param>
    Public Sub Log(text As String, info As Info, loglevel As Integer,
                   <CallerMemberName> Optional memberName As String = "",
                   <CallerFilePath> Optional filename As String = "")

        If loglevel > _loglevel Then Return 'returns if this logline at higher level of detail than global setting

        Dim outline As String = Nothing

        Select Case info
            Case Info.CallingClass
                outline = Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & text
            Case Info.CallingMethod
                outline = memberName.ToUpper & " | " & text
            Case Info.CallingClassAndMethod
                outline = Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & memberName.ToUpper & " | " & text
        End Select

        Write_Line(outline)

    End Sub

    Public Shared Function CallingMethod(<CallerMemberName> Optional memberName As String = "") As String
        Return memberName
    End Function

    ''' <summary>
    ''' Logs the system information for the current user.
    ''' </summary>
    ''' <remarks>A large chunk of this code was borrowed from Ben Baker aka Headkaze.</remarks>
    Private Function Get_System_Info() As Boolean
        Try
            Write_Line("Diagnostics: Begin system enumeraion...")
            Dim query As New Management.ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
            For Each mo As Management.ManagementObject In query.[Get]()
                Dim Total_RAM As [Double] = [Double].Parse(mo("TotalVisibleMemorySize").ToString()) / 1024
                Dim Free_RAM As [Double] = [Double].Parse(mo("FreePhysicalMemory").ToString()) / 1024
                Write_Line("OS: " & mo("Caption").ToString())
                Write_Line("Version: " & mo("Version").ToString())
                Write_Line("Build: " & mo("BuildNumber").ToString())
                Write_Line("Total RAM: " & Math.Ceiling(Total_RAM) & " MB")
                Write_Line("Available RAM: " & Math.Ceiling(Total_RAM - Free_RAM) & " MB")
            Next mo
            query = New Management.ManagementObjectSearcher("SELECT * FROM Win32_Processor")
            For Each obj As Management.ManagementObject In query.[Get]()
                Write_Line("CPU: " & obj("Name").ToString().TrimStart())
            Next obj
            query = New Management.ManagementObjectSearcher("SELECT * FROM Win32_VideoController")
            For Each obj As Management.ManagementObject In query.[Get]()
                Write_Line("Video Card: " & obj("Name").ToString())
                Write_Line("Video Driver: " & obj("DriverVersion").ToString())
                If obj("AdapterRAM") IsNot Nothing Then
                    Dim Video_RAM As Double = Double.Parse(obj("AdapterRAM").ToString()) / 1024 / 1024
                    Write_Line("Video RAM: " & Video_RAM & " MB")
                End If
            Next obj
            query = New Management.ManagementObjectSearcher("SELECT * FROM Win32_SoundDevice")
            For Each obj As Management.ManagementObject In query.[Get]()
                Write_Line("Sound Card: " & obj("Name").ToString())
            Next obj
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v1.0") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 1.0 Installed")
            End If
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v1.1") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 1.1 Installed")
            End If
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v2.0") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 2.0 Installed")
            End If
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v3.0") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 3.0 Installed")
            End If
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v4.0") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 4.0 Installed")
            End If
            If Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\.NETFramework\policy\v4.5") IsNot Nothing Then
                Write_Line(".NET: .NET Framework 4.5 Installed")
            End If
            query.Dispose()
            query = Nothing
            Write_Line("Diagnostics: System enumeraion completed successfully!")
            Return True
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "WRITE ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Attempts to write the system details of the client machine to the log. If fails - returns false
    ''' </summary>
    Public Function LogSystemDetails() As Boolean
        LogSystemDetails = Get_System_Info()
    End Function

    ''' <summary>
    ''' Makes log entry with prefix "ERROR: "
    ''' </summary>
    ''' <param name="text">The text to print to a log line</param>
    Public Sub Err(text As String)
        Write_Line("ERROR: " & text)
    End Sub

    ''' <summary>
    '''  Makes log entry with prefix "ERROR: "
    '''  </summary>
    '''  <param name="text">The text to print to a log line</param>
    ''' <param name="info">Additional info to include in log line.</param>
    Public Sub Err(text As String, info As Info, <CallerMemberName> Optional memberName As String = "",
                   <CallerFilePath> Optional filename As String = "")

        Dim outline As String = Nothing

        Select Case info
            Case Info.CallingClass
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & text
            Case Info.CallingMethod
                outline = "ERROR: " & memberName.ToUpper & " | " & text
            Case Info.CallingClassAndMethod
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & memberName.ToUpper & " | " & text
        End Select

        Write_Line(outline)

    End Sub

    ''' <summary>
    '''   Makes log entry with prefix "ERROR: "
    '''   </summary>
    '''   <param name="text">The text to print to a log line</param>
    '''  <param name="info">Additional info to include in the log line</param>
    ''' <param name="ex">Logs the text and the Exception details (.message and .stacktrace)</param>
    Public Sub Err(text As String, info As Info, ex As System.Exception,
                   <CallerMemberName> Optional memberName As String = "",
                   <CallerFilePath> Optional filename As String = "")

        Dim outline As String = Nothing

        Select Case info
            Case Info.CallingClass
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & text
            Case Info.CallingMethod
                outline = "ERROR: " & memberName.ToUpper & " | " & text
            Case Info.CallingClassAndMethod
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & memberName.ToUpper & " | " & text
        End Select

        Write_Line(outline & vbCr & "Exception: " & ex.Message & vbCr & ex.StackTrace)

    End Sub

    ''' <summary>
    '''    Makes log entry with prefix "ERROR: "
    '''    </summary>
    '''    <param name="text">The text to print to a log line</param>
    '''   <param name="info">Additional info to include in the log line</param>
    '''  <param name="ex">Logs the text and the Exception details (.message and .stacktrace)</param>
    ''' <param name="loglevel">The log level of the line</param>
    Public Sub Err(text As String, info As Info, ex As System.Exception,
                    loglevel As Integer, <CallerMemberName> Optional memberName As String = "",
                   <CallerFilePath> Optional filename As String = "")

        If loglevel > _loglevel Then Return 'returns if this logline at higher level of detail than global setting

        Dim outline As String = Nothing

        Select Case info
            Case Info.CallingClass
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & text
            Case Info.CallingMethod
                outline = "ERROR: " & memberName.ToUpper & " | " & text
            Case Info.CallingClassAndMethod
                outline = "ERROR: " & Path.GetFileNameWithoutExtension(filename).ToUpper & " | " & memberName.ToUpper & " | " & text
        End Select

        Write_Line(outline & vbCr & "Exception: " & ex.Message & vbCr & ex.StackTrace)

    End Sub
End Class

Public Enum LogType
    ''' <summary>
    ''' Erases old log on StartLogging and starts fresh log.
    ''' </summary>
    OverwriteOld
    ''' <summary>
    ''' Appends log to any previous logging in specified file.
    ''' </summary>
    Appended
End Enum

Public Enum Info
    ''' <summary>
    ''' No additional info
    ''' </summary>
    None
    ''' <summary>
    ''' Includes the Class where the log call came from
    ''' </summary>
    CallingClass
    ''' <summary>
    ''' Includes the Method where the log call came from
    ''' </summary>
    CallingMethod
    ''' <summary>
    ''' Includes the Class and Method where the log call came from
    ''' </summary>
    CallingClassAndMethod
End Enum
