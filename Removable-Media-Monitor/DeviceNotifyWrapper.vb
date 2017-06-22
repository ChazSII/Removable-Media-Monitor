Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32.SafeHandles

Friend Class DeviceNotifyWrapper
    Const FILE_FLAG_BACKUP_SEMANTICS As UInteger = &H2000000
    Shared ReadOnly INVALID_HANDLE_VALUE As New IntPtr(-1)

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

#If CONFIG = "Debug" Then
    Public ReadOnly Property NotifyHandle As IntPtr
        Get
            Return mDeviceNotifyHandle
        End Get
    End Property
#End If

    Friend ReadOnly Property CurrentDevice As String
        Get
            Return mCurrentDrive
        End Get
    End Property

    ''' <summary>
    ''' Gets the value indicating whether the query remove event will be fired.
    ''' </summary>
    Friend ReadOnly Property IsQueryHooked As Boolean
        Get
            If mDeviceNotifyHandle = IntPtr.Zero Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets the file stream for file which this class opened on a drive to be notified
    ''' about it's removal. 
    ''' This will be null unless you specified a file to open (DriveDetector opens root directory of the flash drive) 
    ''' </summary>
    Friend ReadOnly Property OpenedFile As FileStream
        Get
            Return mFileOnFlash
        End Get
    End Property


    Friend Sub New(messageHandler As IntPtr, Optional fileToOpen As String = Nothing)
        mFileToOpen = fileToOpen
        mFileOnFlash = Nothing
        mDeviceNotifyHandle = IntPtr.Zero
        mRecipientHandle = messageHandler
        mDirHandle = IntPtr.Zero
        mCurrentDrive = ""
    End Sub


    ''' <summary>
    ''' Hooks specified drive to receive a message when it is being removed.  
    ''' This can be achieved also by setting e.HookQueryRemove to true in your 
    ''' DeviceArrived event handler. 
    ''' By default DriveDetector will open the root directory of the flash drive to obtain notification handle
    ''' from Windows (to learn when the drive is about to be removed). 
    ''' </summary>
    ''' <param name="fileOnDrive">Drive letter or relative path to a file on the drive which should be 
    ''' used to get a handle - required for registering to receive query remove messages.
    ''' If only drive letter is specified (e.g. "D:\", root directory of the drive will be opened.</param>
    ''' <returns>true if hooked ok, false otherwise</returns>
    Friend Function EnableQueryRemove(fileOnDrive As String) As Boolean
        If fileOnDrive Is Nothing OrElse fileOnDrive.Length = 0 Then
            Throw New ArgumentException("Drive path must be supplied to register for Query remove.")
        End If

        If fileOnDrive.Length = 2 AndAlso fileOnDrive(1) = ":" Then
            ' append "\" if only drive letter with ":" was passed in.
            fileOnDrive += "\"
        End If

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
    Friend Sub DisableQueryRemove()
        If mDeviceNotifyHandle <> IntPtr.Zero Then
            RegisterForDeviceChange(False, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' Registers for receiving the query remove message for a given drive.
    ''' We need to open a handle on that drive and register with this handle. 
    ''' Client can specify this file in mFileToOpen or we will open root directory of the drive
    ''' </summary>
    ''' <param name="drive">drive for which to register. </param>
    Private Sub RegisterQuery(drive As String)
        Try
            ' Make sure the path in mFileToOpen contains valid drive
            ' If there is a drive letter in the path, it may be different from the  actual
            ' letter assigned to the drive now. We will cut it off and merge the actual drive 
            ' with the rest of the path.
            If mFileToOpen IsNot Nothing Then
                If mFileToOpen.Contains(":") Then
                    Dim tmp As String = mFileToOpen.Substring(3)
                    Dim root As String = Path.GetPathRoot(drive)

                    mFileToOpen = Path.Combine(root, tmp)
                Else
                    mFileToOpen = Path.Combine(drive, mFileToOpen)
                End If

                mFileOnFlash = New FileStream(mFileToOpen, IO.FileMode.Open)
            End If

            If mFileOnFlash Is Nothing Then
                ' Change 28.10.2007 - Open the root directory 
                RegisterForDeviceChange(drive)
            Else
                ' old version
                RegisterForDeviceChange(True, mFileOnFlash.SafeFileHandle)
            End If

            mCurrentDrive = drive
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' New version which gets the handle automatically for specified directory
    ''' Only for registering! Unregister with the old version of this function...
    ''' </summary>
    ''' <param name="dirPath">e.g. C:\dir</param>
    Private Sub RegisterForDeviceChange(dirPath As String)
        Dim handle As IntPtr = OpenDirectory(dirPath)

        If handle = IntPtr.Zero Then
            mDeviceNotifyHandle = IntPtr.Zero

            Exit Sub
        Else
            ' save handle for closing it when unregistering
            mDirHandle = handle
        End If

        ' Register for handle
        Dim data As New DEV_BROADCAST_HANDLE() With {
            .dbch_devicetype = DeviceTypes.Handle,
            .dbch_reserved = 0,
            .dbch_nameoffset = 0,
            .dbch_handle = handle,
            .dbch_hdevnotify = CType(0, IntPtr)
        }

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
            Dim data As New DEV_BROADCAST_HANDLE() With {
                .dbch_devicetype = DeviceTypes.Handle,
                .dbch_reserved = 0,
                .dbch_nameoffset = 0,
                .dbch_handle = fileHandle.DangerousGetHandle(),
                .dbch_hdevnotify = IntPtr.Zero
            }

            Dim size As Integer = Marshal.SizeOf(data)
            data.dbch_size = size

            Dim buffer As IntPtr = Marshal.AllocHGlobal(size)
            Marshal.StructureToPtr(data, buffer, True)

            mDeviceNotifyHandle = NativeMethods.RegisterDeviceNotification(mRecipientHandle, buffer, 0)
        Else
            ' close the directory handle
            If mDirHandle <> IntPtr.Zero Then
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
    ''' Opens a directory, returns it's handle or zero.
    ''' </summary>
    ''' <param name="dirPath">path to the directory, e.g. "C:\dir"</param>
    ''' <returns>handle to the directory. Close it with CloseHandle().</returns>
    Private Shared Function OpenDirectory(dirPath As String) As IntPtr
        ' open the existing file for reading          
        Dim handle As IntPtr = NativeMethods.CreateFile(dirPath,
                                                        FileAccess.Read,
                                                        FileShare.Read Or FileShare.Write,
                                                        0,
                                                        FileMode.Open,
                                                        FileAttributes.Normal Or FILE_FLAG_BACKUP_SEMANTICS,
                                                        0)

        If handle = INVALID_HANDLE_VALUE Then
            Return IntPtr.Zero
        Else
            Return handle
        End If
    End Function

End Class
