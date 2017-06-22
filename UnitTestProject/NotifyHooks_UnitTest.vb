Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports CS2Soft.RemovableMediaMonitor
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UnitTestProject.Mocking

<TestClass()> Public Class NotifyHooks_UnitTest

    Public Const IMD_SERVICE_REG_KEY As String = "SYSTEM\CurrentControlSet\Services\Imdisk"

    <TestMethod()> Public Sub TestImDisk()
        Try
            'Dim driverPath As String = "imdisk.sys"
            'Dim RegistryService As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(IMD_SERVICE_REG_KEY)

            'If Not IsNothing(RegistryService) AndAlso RegistryService.GetValue("Start") = SERVICE_START_TYPE.AUTO_START Then

            'End If

            'Using ServiceManager As New ServiceManager("Imdisk", "Imdisk Driver", driverPath)
            '    If Not ServiceManager.ServiceExists Then
            '        ServiceManager.CreateWin32Service()
            '    End If

            '    If ServiceManager.StartWin32Service() Then
            '        ServiceManager.RemoveWin32Service()
            '    End If
            'End Using

            Dim devNumber As UInt32 = UInt32.MaxValue
            LTR.IO.ImDisk.ImDiskAPI.CreateDevice(128 * 1024, "Q:", devNumber)

            'Using vDisk As New LTR.IO.ImDisk.ImDiskDevice("Q:", IO.FileAccess.ReadWrite)
            '    Threading.Thread.Sleep(5000)
            'End Using

            For i As Integer = 0 To 5000 Step 1
                Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()

                Exit For
            Next

        Catch ex As Exception
            Dim mtp = ex
            'Finally
            '    Dim driverUnloaded As Boolean = False

            '    Try
            '        'Stop driver services
            '        Using TC_SVC As New ServiceProcess.ServiceController("Imdisk")
            '            For i As Integer = 1 To 10
            '                If Not TC_SVC.Status = ServiceProcess.ServiceControllerStatus.Stopped Then
            '                    TC_SVC.WaitForStatus(ServiceProcess.ServiceControllerStatus.Stopped, New TimeSpan(1000))
            '                Else
            '                    driverUnloaded = True
            '                    Exit For
            '                End If
            '            Next
            '        End Using
            '    Catch ex As Exception
            '        'Servizio inesistente
            '        Dim tmp = ex
            '    End Try
        End Try
    End Sub


    '<TestMethod()> Public Sub TestDriveQueryRemove()
    '    Dim eventRaised As Boolean = False
    '    Dim errorRaised As Boolean = False

    '    Dim device As New MockedDevice(LetterFlags.O)


    '    'Dim tmpFile As String = IO.Path.GetTempFileName()
    '    Using vhdDev As New Medo.IO.VirtualDisk("C:\Users\cshrawder\Desktop\test2.vhd")
    '        vhdDev.Open()

    '        Using devDetector As New DeviceChangeDetector()
    '            Dim deviceArrived As Boolean = False
    '            AddHandler devDetector.DeviceArrived, Sub(sender, e)
    '                                                      e.HookQueryRemove = True
    '                                                      deviceArrived = True
    '                                                  End Sub

    '            AddHandler devDetector.QueryRemove, Sub(sender, e)
    '                                                    Dim eventLetter As String = e.Drive(0)
    '                                                    Dim deviceLetter As String = [Enum].GetName(GetType(LetterFlags),
    '                                                                                             device.DriveLetter)
    '                                                    If eventLetter = deviceLetter Then
    '                                                        eventRaised = True
    '                                                    End If
    '                                                End Sub

    '            vhdDev.Attach(Medo.IO.VirtualDiskAttachOptions.None)

    '            For delay As Integer = 0 To 50
    '                System.Windows.Forms.Application.DoEvents()

    '                If deviceArrived Then Exit For

    '                Threading.Thread.Sleep(100)
    '            Next

    '            vhdDev.Detach()

    '            For delay As Integer = 0 To 50
    '                System.Windows.Forms.Application.DoEvents()

    '                If eventRaised Then Exit For

    '                Threading.Thread.Sleep(100)
    '            Next

    '            vhdDev.Close()

    '        End Using
    '    End Using

    '    Assert.IsTrue(eventRaised)
    '    Assert.IsFalse(errorRaised)
    'End Sub

End Class


'Friend Class ServiceManager
'    Implements IDisposable

'    ' Service Constants
'    Private Const SC_MANAGER_ALL_ACCESS As Integer = &HF003F 'Service Control Manager object specific access types
'    Private Const SC_MANAGER_CREATE_SERVICE As Integer = &H2 'Service Control Manager create service access
'    Private Const SERVICE_ALL_ACCESS As Integer = &HF01FF    'Service object specific access type
'    Private Const SERVICE_KERNEL_DRIVER As Integer = &H1     'Service Type
'    Private Const SERVICE_ERROR_NORMAL As Integer = &H1      'Error control type
'    Private Const GENERIC_WRITE As Integer = &H40000000
'    Private Const DELETE As Integer = &H10000

'    ' Service Handles
'    Private Property hSCManager As IntPtr = IntPtr.Zero
'    Private Property hService As IntPtr = IntPtr.Zero
'    Private Property Service As String
'    Private Property Display As String
'    Private Property DriverPath As String

'    ' Private properties
'    Private Property _ServiceExists As Boolean = False
'    Private Property _LastWin32Error As Integer = 0

'    ' Public properties
'    Public ReadOnly Property ServiceExists As Boolean
'        Get
'            Return _ServiceExists
'        End Get
'    End Property

'    Public ReadOnly Property LastWin32Error As Integer
'        Get
'            Return _LastWin32Error
'        End Get
'    End Property



'    Protected Friend Sub New(ByVal ServiceName As String, ByVal DisplayName As String, ByVal DriverLocation As String)
'        Service = ServiceName
'        Display = DisplayName
'        DriverPath = DriverLocation

'        hSCManager = OpenSCManager(Nothing, Nothing, SC_MANAGER_ALL_ACCESS)

'        If hSCManager = IntPtr.Zero Then
'            _LastWin32Error = Marshal.GetLastWin32Error()
'            Dim tmp = "Couldn't open service manager"
'        Else
'            hService = OpenService(hSCManager, Service, SERVICE_ALL_ACCESS)
'        End If

'        If hService.Equals(IntPtr.Zero) Then
'            _ServiceExists = False
'        Else
'            _ServiceExists = True
'        End If
'    End Sub

'    Protected Friend Function CreateWin32Service() As Boolean
'        hService = CreateService(hSCManager, Service, Display,
'                                            SERVICE_ALL_ACCESS, SERVICE_KERNEL_DRIVER,
'                                            SERVICE_START_TYPE.DEMAND_START, SERVICE_ERROR_NORMAL, DriverPath,
'                                            Nothing, 0, Nothing, Nothing, Nothing)

'        If hService.Equals(IntPtr.Zero) Then
'            _LastWin32Error = Marshal.GetLastWin32Error()
'            _ServiceExists = False

'            Dim tmp = "Couldn't create service"
'        Else
'            _ServiceExists = True

'            Return True
'        End If

'        Return False
'    End Function

'    Protected Friend Function StartWin32Service() As Boolean
'        If ServiceExists Then
'            'now trying to start the service
'            Dim intReturnVal As Integer = StartService(hService, 0, Nothing)

'            ' If the value i is zero, then there was an error starting the service.
'            ' note: error may arise if the service is already running or some other problem.
'            If intReturnVal = 0 Then
'                _LastWin32Error = Marshal.GetLastWin32Error()
'                StartWin32Service = False
'            Else
'                StartWin32Service = True
'            End If
'        End If

'        Return StartWin32Service
'    End Function

'    Protected Friend Function StopWin32Service() As Boolean
'        Using TC_SVC As New ServiceProcess.ServiceController("TrueCrypt")
'            For i As Integer = 1 To 10
'                If Not TC_SVC.Status = ServiceProcess.ServiceControllerStatus.Stopped Then
'                    TC_SVC.WaitForStatus(ServiceProcess.ServiceControllerStatus.Stopped, New TimeSpan(1000))
'                Else
'                    Return True
'                End If
'            Next
'        End Using

'        Return False
'        'If _ServiceExists Then
'        '    'now trying to start the service
'        '    Dim intReturnVal As Integer = StopService(hService, 0, Nothing)

'        '    ' If the value i is zero, then there was an error starting the service.
'        '    ' note: error may arise if the service is already running or some other problem.
'        '    If intReturnVal = 0 Then
'        '        StopWin32Service = True
'        '    Else
'        '        _LastWin32Error = Marshal.GetLastWin32Error()
'        '        StopWin32Service = False
'        '    End If
'        'End If

'        'Return StopWin32Service
'    End Function

'    Protected Friend Function RemoveWin32Service() As Boolean
'        If ServiceExists Then
'            'now trying to start the service
'            Dim intReturnVal As Integer = DeleteService(hService)

'            ' If the value i is zero, then there was an error starting the service.
'            ' note: error may arise if the service is already running or some other problem.
'            If intReturnVal = 0 Then
'                RemoveWin32Service = True
'                CloseServiceHandle(hService)
'            Else
'                _LastWin32Error = Marshal.GetLastWin32Error()
'                RemoveWin32Service = False
'            End If
'        End If

'        Return StartWin32Service()
'    End Function

'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' To detect redundant calls

'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not Me.disposedValue Then
'            If disposing Then
'                ' TODO: dispose managed state (managed objects).
'            End If

'            CloseServiceHandle(hService)

'            Try
'                CloseHandle(hSCManager)
'            Catch ex As Exception
'                _LastWin32Error = Marshal.GetLastWin32Error()
'                Dim tmp = ""
'            End Try

'            hSCManager = IntPtr.Zero
'            hService = IntPtr.Zero

'            _ServiceExists = Nothing
'            _LastWin32Error = Nothing


'            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'            ' TODO: set large fields to null.
'        End If
'        Me.disposedValue = True
'    End Sub

'    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
'    'Protected Overrides Sub Finalize()
'    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub

'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'        Dispose(True)
'        GC.SuppressFinalize(Me)
'    End Sub
'#End Region

'End Class


'Friend Module NativeMethods

'    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
'    Public Function CreateFile(lpFileName As String, <MarshalAs(UnmanagedType.U4)> dwDesiredAccess As FileAccess, <MarshalAs(UnmanagedType.U4)> dwShareMode As FileShare, lpSecurityAttributes As IntPtr, <MarshalAs(UnmanagedType.U4)> dwCreationDisposition As FileMode, <MarshalAs(UnmanagedType.U4)> dwFlagsAndAttributes As FileAttributes, hTemplateFile As IntPtr) As IntPtr
'    End Function

'    <DllImport("kernel32.dll", SetLastError:=True)>
'    Public Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
'    End Function

'    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
'    Public Function OpenSCManager(ByVal machineName As String, ByVal databaseName As String, ByVal desiredAccess As Int32) As IntPtr
'    End Function

'    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
'    Public Function CreateService(ByVal hSCManager As IntPtr, ByVal serviceName As String, ByVal displayName As String, ByVal desiredAccess As Int32, ByVal serviceType As Int32, ByVal startType As Int32, ByVal errorcontrol As Int32, ByVal binaryPathName As String, ByVal loadOrderGroup As String, ByVal TagBY As Int32, ByVal dependencides As String, ByVal serviceStartName As String, ByVal password As String) As IntPtr
'    End Function

'    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
'    Public Function OpenService(ByVal hSCManager As IntPtr, ByVal lpServiceName As String, ByVal dwDesiredAccess As Int32) As IntPtr
'    End Function

'    <DllImport("advapi32", SetLastError:=True)>
'    Public Function StartService(hService As IntPtr, dwNumServiceArgs As Integer, lpServiceArgVectors As String()) As <MarshalAs(UnmanagedType.Bool)> Boolean
'    End Function

'    <DllImport("advapi32.dll", SetLastError:=True)>
'    Public Function ControlService(ByVal hService As IntPtr, ByVal dwControl As SERVICE_CONTROL, ByRef lpServiceStatus As SERVICE_STATE) As Boolean
'    End Function

'    <DllImport("advapi32.dll", SetLastError:=True)>
'    Public Function CloseServiceHandle(ByVal serviceHandle As IntPtr) As Boolean
'    End Function

'    <DllImport("advapi32.dll", SetLastError:=True)>
'    Public Function DeleteService(ByVal hService As IntPtr) As Boolean
'    End Function
'End Module

'' Service Start Type
'Friend Enum SERVICE_START_TYPE As Integer
'    SYSTEM_START = &H1
'    AUTO_START = &H2
'    DEMAND_START = &H3
'End Enum

'Friend Enum SERVICE_CONTROL As Integer
'    [STOP] = &H1
'    PAUSE = &H2
'    [CONTINUE] = &H3
'    INTERROGATE = &H4
'    SHUTDOWN = &H5
'    PARAMCHANGE = &H6
'    NETBINDADD = &H7
'    NETBINDREMOVE = &H8
'    NETBINDENABLE = &H9
'    NETBINDDISABLE = &HA
'    DEVICEEVENT = &HB
'    HARDWAREPROFILECHANGE = &HC
'    POWEREVENT = &HD
'    SESSIONCHANGE = &HE
'End Enum

'' Service States
'Friend Enum SERVICE_STATE As Integer
'    SERVICE_STOPPED = &H1
'    SERVICE_START_PENDING = &H2
'    SERVICE_STOP_PENDING = &H3
'    SERVICE_RUNNING = &H4
'    SERVICE_CONTINUE_PENDING = &H5
'    SERVICE_PAUSE_PENDING = &H6
'    SERVICE_PAUSED = &H7
'End Enum

'Friend Enum SERVICE_ACCEPT As Integer
'    [STOP] = &H1
'    PAUSE_CONTINUE = &H2
'    SHUTDOWN = &H4
'    PARAMCHANGE = &H8
'    NETBINDCHANGE = &H10
'    HARDWAREPROFILECHANGE = &H20
'    POWEREVENT = &H40
'    SESSIONCHANGE = &H80
'End Enum