''==================================================================''
''                                                                  ''
''      Copied and modified from Jan Dolinay's Drive Detector       ''
''          - http://www.codeproject.com/Articles/18062/            ''
''                                                                  ''
''  Translated from C# to VB via Snippet Converter and manually     ''
''  - http://codeconverter.sharpdevelop.net/SnippetConverter.aspx   ''
''                                                                  ''
''==================================================================''
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms ' required for Message
Imports System.Runtime.InteropServices ' required for Marshal
Imports System.IO
Imports Microsoft.Win32.SafeHandles
' DriveDetector - rev. 1, Oct. 31 2007

''' <summary>
''' Hidden Form which we use to receive Windows messages about flash drives
''' </summary>
Friend Class DetectorForm
    Inherits Form
    Private label1 As Label
    Private mDetector As DriveDetector = Nothing

    ''' <summary>
    ''' Set up the hidden form. 
    ''' </summary>
    ''' <param name="detector">DriveDetector object which will receive notification about USB drives, see WndProc</param>
    Public Sub New(detector As DriveDetector)
        mDetector = detector
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ShowInTaskbar = False
        Me.ShowIcon = False
        Me.FormBorderStyle = FormBorderStyle.None
        AddHandler Me.Load, New System.EventHandler(AddressOf Me.Load_Form)
        AddHandler Me.Activated, New EventHandler(AddressOf Me.Form_Activated)
    End Sub

    Private Sub Load_Form(sender As Object, e As EventArgs)
        ' We don't really need this, just to display the label in designer ...
        InitializeComponent()

        ' Create really small form, invisible anyway.
        Me.Size = New System.Drawing.Size(5, 5)
    End Sub

    Private Sub Form_Activated(sender As Object, e As EventArgs)
        Me.Visible = False
    End Sub

    ''' <summary>
    ''' This function receives all the windows messages for this window (form).
    ''' We call the DriveDetector from here so that is can pick up the messages about
    ''' drives arrived and removed.
    ''' </summary>
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If mDetector IsNot Nothing Then
            mDetector.WndProc(m)
        End If
    End Sub

    Private Sub InitializeComponent()
        Me.label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        ' 
        ' label1
        ' 
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(13, 30)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(314, 13)
        Me.label1.TabIndex = 0
        Me.label1.Text = "This is invisible form. To see DriveDetector code click View Code"
        ' 
        ' DetectorForm
        ' 
        Me.ClientSize = New System.Drawing.Size(360, 80)
        Me.Controls.Add(Me.label1)
        Me.Name = "DetectorForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class

' Delegate for event handler to handle the device events 
Public Delegate Sub DriveDetectorEventHandler(sender As [Object], e As DriveDetectorEventArgs)

''' <summary>
''' Our class for passing in custom arguments to our event handlers 
''' </summary>
Public Class DriveDetectorEventArgs
    Inherits EventArgs

    Public Sub New()
        Cancel = False
        Drive = ""
        HookQueryRemove = False
    End Sub

    ''' <summary>
    ''' Get/Set the value indicating that the event should be cancelled 
    ''' Only in QueryRemove handler.
    ''' </summary>
    Public Cancel As Boolean

    ''' <summary>
    ''' Drive letter for the device which caused this event 
    ''' </summary>
    Public Drive As String

    ''' <summary>
    ''' Set to true in your DeviceArrived event handler if you wish to receive the 
    ''' QueryRemove event for this drive. 
    ''' </summary>
    Public HookQueryRemove As Boolean

End Class


''' <summary>
''' Detects insertion or removal of removable drives.
''' Use it in 1 or 2 steps:
''' 1) Create instance of this class in your project and add handlers for the
''' DeviceArrived, DeviceRemoved and QueryRemove events.
''' AND (if you do not want drive detector to creaate a hidden form))
''' 2) Override WndProc in your form and call DriveDetector's WndProc from there. 
''' If you do not want to do step 2, just use the DriveDetector constructor without arguments and
''' it will create its own invisible form to receive messages from Windows.
''' </summary>
Public Class DriveDetector
    Implements IDisposable

    ''' <summary>
    ''' Events signalized to the client app.
    ''' Add handlers for these events in your form to be notified of removable device events 
    ''' </summary>
    Public Event DeviceArrived As DriveDetectorEventHandler
    Public Event DeviceRemoved As DriveDetectorEventHandler
    Public Event QueryRemove As DriveDetectorEventHandler

    ''' <summary>
    ''' The easiest way to use DriveDetector. 
    ''' It will create hidden form for processing Windows messages about USB drives
    ''' You do not need to override WndProc in your form.
    ''' </summary>
    Public Sub New()
        Dim frm As New DetectorForm(Me)
        frm.Show()
        ' will be hidden immediatelly
        Init(frm, Nothing)
    End Sub

    ''' <summary>
    ''' Alternate constructor.
    ''' Pass in your Form and DriveDetector will not create hidden form.
    ''' </summary>
    ''' <param name="control">object which will receive Windows messages. 
    ''' Pass "this" as this argument from your form class.</param>
    Public Sub New(control As Control)
        Init(control, Nothing)
    End Sub

    ''' <summary>
    ''' Consructs DriveDetector object setting also path to file which should be opened
    ''' when registering for query remove.  
    ''' </summary>
    '''<param name="control">object which will receive Windows messages. 
    ''' Pass "this" as this argument from your form class.</param>
    ''' <param name="FileToOpen">Optional. Name of a file on the removable drive which should be opened. 
    ''' If null, root directory of the drive will be opened. Opening a file is needed for us 
    ''' to be able to register for the query remove message. TIP: For files use relative path without drive letter.
    ''' e.g. "SomeFolder\file_on_flash.txt"</param>
    Public Sub New(control As Control, FileToOpen As String)
        Init(control, FileToOpen)
    End Sub

    ''' <summary>
    ''' init the DriveDetector object
    ''' </summary>
    Private Sub Init(control As Control, fileToOpen As String)
        mFileToOpen = fileToOpen
        mFileOnFlash = Nothing
        mDeviceNotifyHandle = IntPtr.Zero
        mRecipientHandle = control.Handle
        mDirHandle = IntPtr.Zero
        ' handle to the root directory of the flash drive which we open 
        mCurrentDrive = ""
    End Sub

    ''' <summary>
    ''' Gets the value indicating whether the query remove event will be fired.
    ''' </summary>
    Public ReadOnly Property IsQueryHooked() As Boolean
        Get
            If mDeviceNotifyHandle = IntPtr.Zero Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets letter of drive which is currently hooked. Empty string if none.
    ''' See also IsQueryHooked.
    ''' </summary>
    Public ReadOnly Property HookedDrive() As String
        Get
            Return mCurrentDrive
        End Get
    End Property

    ''' <summary>
    ''' Gets the file stream for file which this class opened on a drive to be notified
    ''' about it's removal. 
    ''' This will be null unless you specified a file to open (DriveDetector opens root directory of the flash drive) 
    ''' </summary>
    Public ReadOnly Property OpenedFile() As FileStream
        Get
            Return mFileOnFlash
        End Get
    End Property

    ''' <summary>
    ''' Hooks specified drive to receive a message when it is being removed.  
    ''' This can be achieved also by setting e.HookQueryRemove to true in your 
    ''' DeviceArrived event handler. 
    ''' By default DriveDetector will open the root directory of the flash drive to obtain notification handle
    ''' from Windows (to learn when the drive is about to be removed). 
    ''' </summary>
    ''' <param name="fileOnDrive">Drive letter or relative path to a file on the drive which should be 
    ''' used to get a handle - required for registering to receive query remove messages.
    ''' If only drive letter is specified (e.g. "D:\\", root directory of the drive will be opened.</param>
    ''' <returns>true if hooked ok, false otherwise</returns>
    Public Function EnableQueryRemove(fileOnDrive As String) As Boolean
        If fileOnDrive Is Nothing OrElse fileOnDrive.Length = 0 Then
            Throw New ArgumentException("Drive path must be supplied to register for Query remove.")
        End If

        If fileOnDrive.Length = 2 AndAlso fileOnDrive(1) = ":"c Then
            fileOnDrive += "\"c
        End If
        ' append "\\" if only drive letter with ":" was passed in.
        If mDeviceNotifyHandle <> IntPtr.Zero Then
            ' Unregister first...
            RegisterForDeviceChange(False, Nothing)
        End If

        If Path.GetFileName(fileOnDrive).Length = 0 OrElse Not File.Exists(fileOnDrive) Then
            mFileToOpen = Nothing
        Else
            ' use root directory...
            mFileToOpen = fileOnDrive
        End If

        RegisterQuery(Path.GetPathRoot(fileOnDrive))
        If mDeviceNotifyHandle = IntPtr.Zero Then
            Return False
        End If
        ' failed to register
        Return True
    End Function

    ''' <summary>
    ''' Unhooks any currently hooked drive so that the query remove 
    ''' message is not generated for it.
    ''' </summary>
    Public Sub DisableQueryRemove()
        If mDeviceNotifyHandle <> IntPtr.Zero Then
            RegisterForDeviceChange(False, Nothing)
        End If
    End Sub


#Region "WindowProc"
    ''' <summary>
    ''' Message handler which must be called from client form.
    ''' Processes Windows messages and calls event handlers. 
    ''' </summary>
    ''' <param name="m"></param>
    Public Sub WndProc(ByRef m As Message)
        Dim devType As Integer
        Dim driveLetter As Char

        If m.Msg = WM_DEVICECHANGE Then
            ' WM_DEVICECHANGE can have several meanings depending on the WParam value...

            Select Case m.WParam.ToInt32()
                ' New device has just arrived
                Case DBT_DEVICEARRIVAL
                    devType = Marshal.ReadInt32(m.LParam, 4)

                    If devType = DBT_DEVTYP_VOLUME Then
                        Dim vol As DEV_BROADCAST_VOLUME = CType(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)

                        ' Get the drive letter 
                        driveLetter = DriveMaskToLetter(vol.dbcv_unitmask)

                        ' Call the client event handler
                        Dim e As New DriveDetectorEventArgs() With {.Drive = driveLetter & ":\"}
                        RaiseEvent DeviceArrived(Me, e)

                        ' Register for query remove if requested
                        If e.HookQueryRemove Then
                            ' If something is already hooked, unhook it now
                            If mDeviceNotifyHandle <> IntPtr.Zero Then
                                RegisterForDeviceChange(False, Nothing)
                            End If

                            RegisterQuery(driveLetter & ":\")
                        End If
                    End If
                Case DBT_DEVICEQUERYREMOVE
                    ' Device is about to be removed
                    ' Any application can cancel the removal

                    devType = Marshal.ReadInt32(m.LParam, 4)
                    If devType = DBT_DEVTYP_HANDLE Then
                        ' TODO: we could get the handle for which this message is sent 
                        ' from vol.dbch_handle and compare it against a list of handles for 
                        ' which we have registered the query remove message (?)                                                 
                        'DEV_BROADCAST_HANDLE vol;
                        'vol = (DEV_BROADCAST_HANDLE)
                        '   Marshal.PtrToStructure(m.LParam, typeof(DEV_BROADCAST_HANDLE));
                        ' if ( vol.dbch_handle ....



                        ' Call the event handler in client
                        Dim e As New DriveDetectorEventArgs() With {.Drive = mCurrentDrive}
                        RaiseEvent QueryRemove(Me, e)

                        ' If the client wants to cancel, let Windows know
                        If e.Cancel Then
                            m.Result = CType(BROADCAST_QUERY_DENY, IntPtr)
                        Else
                            ' Change 28.10.2007: Unregister the notification, this will
                            ' close the handle to file or root directory also. 
                            ' We have to close it anyway to allow the removal so
                            ' even if some other app cancels the removal we would not know about it...                                    
                            ' will also close the mFileOnFlash
                            RegisterForDeviceChange(False, Nothing)

                        End If
                    End If
                Case DBT_DEVICEREMOVECOMPLETE
                    ' Device has been removed

                    devType = Marshal.ReadInt32(m.LParam, 4)
                    If devType = DBT_DEVTYP_VOLUME Then
                        Dim vol As DEV_BROADCAST_VOLUME = CType(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                        driveLetter = DriveMaskToLetter(vol.dbcv_unitmask)

                        ' Call the client event handler
                        Dim e As New DriveDetectorEventArgs() With {.Drive = driveLetter & ":\"}
                        RaiseEvent DeviceRemoved(Me, e)

                        ' TODO: we could unregister the notify handle here if we knew it is the
                        ' right drive which has been just removed
                        'RegisterForDeviceChange(false, null);
                    End If
            End Select
        End If

    End Sub
#End Region


#Region "Private Area"

    ''' <summary>
    ''' New: 28.10.2007 - handle to root directory of flash drive which is opened
    ''' for device notification
    ''' </summary>
    Private mDirHandle As IntPtr = IntPtr.Zero

    ''' <summary>
    ''' Class which contains also handle to the file opened on the flash drive
    ''' </summary>
    Private mFileOnFlash As FileStream = Nothing

    ''' <summary>
    ''' Name of the file to try to open on the removable drive for query remove registration
    ''' </summary>
    Private mFileToOpen As String

    ''' <summary>
    ''' Handle to file which we keep opened on the drive if query remove message is required by the client
    ''' </summary>       
    Private mDeviceNotifyHandle As IntPtr

    ''' <summary>
    ''' Handle of the window which receives messages from Windows. This will be a form.
    ''' </summary>
    Private mRecipientHandle As IntPtr

    ''' <summary>
    ''' Drive which is currently hooked for query remove
    ''' </summary>
    Private mCurrentDrive As String


    ' Win32 constants
    Private Const DBT_DEVTYP_DEVICEINTERFACE As Integer = 5
    Private Const DBT_DEVTYP_HANDLE As Integer = 6
    Private Const BROADCAST_QUERY_DENY As Integer = &H424D5144
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    ' system detected a new device
    Private Const DBT_DEVICEQUERYREMOVE As Integer = &H8001
    ' Preparing to remove (any program can disable the removal)
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    ' removed 
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2
    ' drive type is logical volume
    ''' <summary>
    ''' Registers for receiving the query remove message for a given drive.
    ''' We need to open a handle on that drive and register with this handle. 
    ''' Client can specify this file in mFileToOpen or we will open root directory of the drive
    ''' </summary>
    ''' <param name="drive">drive for which to register. </param>
    Private Sub RegisterQuery(drive As String)
        Dim register As Boolean = True

        ' Change 28.10.2007 - Open the root directory if no file specified - leave mFileToOpen null 
        ' If client gave us no file, let's pick one on the drive... 
        'mFileToOpen = GetAnyFile(drive);
        'if (mFileToOpen.Length == 0)
        '    return;     // no file found on the flash drive                
        If mFileToOpen Is Nothing Then
        Else
            ' Make sure the path in mFileToOpen contains valid drive
            ' If there is a drive letter in the path, it may be different from the  actual
            ' letter assigned to the drive now. We will cut it off and merge the actual drive 
            ' with the rest of the path.
            If mFileToOpen.Contains(":") Then
                Dim tmp As String = mFileToOpen.Substring(3)
                Dim root As String = Path.GetPathRoot(drive)
                mFileToOpen = Path.Combine(root, tmp)
            Else
                mFileToOpen = Path.Combine(drive, mFileToOpen)
            End If
        End If


        Try
            'mFileOnFlash = new FileStream(mFileToOpen, FileMode.Open);
            ' Change 28.10.2007 - Open the root directory 
            If mFileToOpen Is Nothing Then
                ' open root directory
                mFileOnFlash = Nothing
            Else
                mFileOnFlash = New FileStream(mFileToOpen, FileMode.Open)
            End If
        Catch generatedExceptionName As Exception
            ' just do not register if the file could not be opened
            register = False
        End Try


        If register Then
            'RegisterForDeviceChange(true, mFileOnFlash.SafeFileHandle);
            'mCurrentDrive = drive;
            ' Change 28.10.2007 - Open the root directory 
            If mFileOnFlash Is Nothing Then
                RegisterForDeviceChange(drive)
            Else
                ' old version
                RegisterForDeviceChange(True, mFileOnFlash.SafeFileHandle)
            End If

            mCurrentDrive = drive
        End If


    End Sub


    ''' <summary>
    ''' New version which gets the handle automatically for specified directory
    ''' Only for registering! Unregister with the old version of this function...
    ''' </summary>
    ''' <param name="dirPath">e.g. C:\\dir</param>
    Private Sub RegisterForDeviceChange(dirPath As String)
        Dim handle As IntPtr = NativeMethods.OpenDirectory(dirPath)
        If handle = IntPtr.Zero Then
            mDeviceNotifyHandle = IntPtr.Zero
            Return
        Else
            mDirHandle = handle
        End If
        ' save handle for closing it when unregistering
        ' Register for handle
        Dim data As New DEV_BROADCAST_HANDLE()
        data.dbch_devicetype = DBT_DEVTYP_HANDLE
        data.dbch_reserved = 0
        data.dbch_nameoffset = 0
        'data.dbch_data = null;
        'data.dbch_eventguid = 0;
        data.dbch_handle = handle
        data.dbch_hdevnotify = CType(0, IntPtr)
        Dim size As Integer = Marshal.SizeOf(data)
        data.dbch_size = size
        Dim buffer As IntPtr = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(data, buffer, True)

        mDeviceNotifyHandle = NativeMethods.RegisterDeviceNotification(mRecipientHandle, buffer, 0)

    End Sub

    ''' <summary>
    ''' Registers to be notified when the volume is about to be removed
    ''' This is requierd if you want to get the QUERY REMOVE messages
    ''' </summary>
    ''' <param name="register">true to register, false to unregister</param>
    ''' <param name="fileHandle">handle of a file opened on the removable drive</param>
    Private Sub RegisterForDeviceChange(register As Boolean, fileHandle As SafeFileHandle)
        If register Then
            ' Register for handle
            Dim data As New DEV_BROADCAST_HANDLE()
            data.dbch_devicetype = DBT_DEVTYP_HANDLE
            data.dbch_reserved = 0
            data.dbch_nameoffset = 0
            'data.dbch_data = null;
            'data.dbch_eventguid = 0;
            data.dbch_handle = fileHandle.DangerousGetHandle()
            'Marshal. fileHandle; 
            data.dbch_hdevnotify = CType(0, IntPtr)
            Dim size As Integer = Marshal.SizeOf(data)
            data.dbch_size = size
            Dim buffer As IntPtr = Marshal.AllocHGlobal(size)
            Marshal.StructureToPtr(data, buffer, True)

            mDeviceNotifyHandle = NativeMethods.RegisterDeviceNotification(mRecipientHandle, buffer, 0)
        Else
            ' close the directory handle
            If mDirHandle <> IntPtr.Zero Then
                '    string er = Marshal.GetLastWin32Error().ToString();
                NativeMethods.CloseDirectoryHandle(mDirHandle)
            End If

            ' unregister
            If mDeviceNotifyHandle <> IntPtr.Zero Then
                NativeMethods.UnregisterDeviceNotification(mDeviceNotifyHandle)
            End If


            mDeviceNotifyHandle = IntPtr.Zero
            mDirHandle = IntPtr.Zero

            mCurrentDrive = ""
            If mFileOnFlash IsNot Nothing Then
                mFileOnFlash.Close()
                mFileOnFlash = Nothing
            End If
        End If

    End Sub

    ''' <summary>
    ''' Gets drive letter from a bit mask where bit 0 = A, bit 1 = B etc.
    ''' There can actually be more than one drive in the mask but we 
    ''' just use the last one in this case.
    ''' </summary>
    ''' <param name="mask"></param>
    ''' <returns></returns>
    Private Shared Function DriveMaskToLetter(mask As Integer) As Char
        Dim letter As Char
        Dim drives As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ' 1 = A
        ' 2 = B
        ' 4 = C...
        Dim cnt As Integer = 0
        Dim pom As Integer = mask \ 2
        While pom <> 0
            ' while there is any bit set in the mask
            ' shift it to the righ...                
            pom = pom \ 2
            cnt += 1
        End While

        If cnt < drives.Length Then
            letter = drives(cnt)
        Else
            letter = "?"c
        End If

        Return letter
    End Function

    ' 28.10.2007 - no longer needed
    '        /// <summary>
    '        /// Searches for any file in a given path and returns its full path
    '        /// </summary>
    '        /// <param name="drive">drive to search</param>
    '        /// <returns>path of the file or empty string</returns>
    '        private string GetAnyFile(string drive)
    '        {
    '            string file = "";
    '            // First try files in the root
    '            string[] files = Directory.GetFiles(drive);
    '            if (files.Length == 0)
    '            {
    '                // if no file in the root, search whole drive
    '                files = Directory.GetFiles(drive, "*.*", SearchOption.AllDirectories);
    '            }
    '                
    '            if (files.Length > 0)
    '                file = files[0];        // get the first file
    '
    '            // return empty string if no file found
    '            return file;
    '        }

#End Region

#Region "Native Win32 API"
    ''' <summary>
    ''' WinAPI functions
    ''' </summary>        
    Private Class NativeMethods
        '   HDEVNOTIFY RegisterDeviceNotification(HANDLE hRecipient,LPVOID NotificationFilter,DWORD Flags);
        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Public Shared Function RegisterDeviceNotification(hRecipient As IntPtr, NotificationFilter As IntPtr, Flags As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Public Shared Function UnregisterDeviceNotification(hHandle As IntPtr) As UInteger
        End Function

        '
        ' CreateFile  - MSDN
        Const OPEN_EXISTING As UInteger = 3
        Const FILE_SHARE_READ As UInteger = &H1
        Const FILE_SHARE_WRITE As UInteger = &H2
        Const FILE_ATTRIBUTE_NORMAL As UInteger = 128
        Const FILE_FLAG_BACKUP_SEMANTICS As UInteger = &H2000000
        Shared ReadOnly INVALID_HANDLE_VALUE As New IntPtr(-1)


        ' should be "static extern unsafe"
        ' file name
        ' access mode
        ' share mode
        ' Security Attributes
        ' how to create
        ' file attributes
        <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Public Shared Function CreateFile(lpFileName As String, <MarshalAs(UnmanagedType.U4)> dwDesiredAccess As FileAccess, <MarshalAs(UnmanagedType.U4)> dwShareMode As FileShare, lpSecurityAttributes As IntPtr, <MarshalAs(UnmanagedType.U4)> dwCreationDisposition As FileMode, <MarshalAs(UnmanagedType.U4)> dwFlagsAndAttributes As FileAttributes, hTemplateFile As IntPtr) As IntPtr
        End Function


        <DllImport("kernel32", SetLastError:=True)> _
        Private Shared Function CloseHandle(hObject As IntPtr) As Boolean
            ' handle to object
        End Function

        ''' <summary>
        ''' Opens a directory, returns it's handle or zero.
        ''' </summary>
        ''' <param name="dirPath">path to the directory, e.g. "C:\\dir"</param>
        ''' <returns>handle to the directory. Close it with CloseHandle().</returns>
        Public Shared Function OpenDirectory(dirPath As String) As IntPtr
            ' open the existing file for reading          
            Dim handle As IntPtr = CreateFile(dirPath, FileAccess.Read, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_ATTRIBUTE_NORMAL, _
                0)

            If handle = INVALID_HANDLE_VALUE Then
                Return IntPtr.Zero
            Else
                Return handle
            End If
        End Function


        Public Shared Function CloseDirectoryHandle(handle As IntPtr) As Boolean
            Return CloseHandle(handle)
        End Function
    End Class


    ' Structure with information for RegisterDeviceNotification.
    <StructLayout(LayoutKind.Sequential)> _
    Friend Structure DEV_BROADCAST_HANDLE
        Friend dbch_size As Integer
        Friend dbch_devicetype As Integer
        Friend dbch_reserved As Integer
        Friend dbch_handle As IntPtr
        Friend dbch_hdevnotify As IntPtr
        Friend dbch_eventguid As Guid
        Friend dbch_nameoffset As Long
        'public byte[] dbch_data[1]; // = new byte[1];
        Friend dbch_data As Byte
        Friend dbch_data1 As Byte
    End Structure

    ' Struct for parameters of the WM_DEVICECHANGE message
    <StructLayout(LayoutKind.Sequential)> _
    Friend Structure DEV_BROADCAST_VOLUME
        Friend dbcv_size As Integer
        Friend dbcv_devicetype As Integer
        Friend dbcv_reserved As Integer
        Friend dbcv_unitmask As Integer
    End Structure
#End Region



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).

            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.

            ' Unregister and close the file we may have opened on the removable drive. 
            RegisterForDeviceChange(False, Nothing)

        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
