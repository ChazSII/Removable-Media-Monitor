Imports System.Runtime.InteropServices

Friend Class NativeMethods
    '   HDEVNOTIFY RegisterDeviceNotification(HANDLE hRecipient,LPVOID NotificationFilter,DWORD Flags);
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function RegisterDeviceNotification(hRecipient As IntPtr, NotificationFilter As IntPtr, Flags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Friend Shared Function UnregisterDeviceNotification(hHandle As IntPtr) As UInteger
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function CreateFile(lpFileName As String, <MarshalAs(UnmanagedType.U4)> dwDesiredAccess As IO.FileAccess, <MarshalAs(UnmanagedType.U4)> dwShareMode As IO.FileShare, lpSecurityAttributes As IntPtr, <MarshalAs(UnmanagedType.U4)> dwCreationDisposition As IO.FileMode, <MarshalAs(UnmanagedType.U4)> dwFlagsAndAttributes As IO.FileAttributes, hTemplateFile As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(hObject As IntPtr) As Boolean
    End Function


    Friend Shared Function CloseDirectoryHandle(handle As IntPtr) As Boolean
        Return CloseHandle(handle)
    End Function

    'Friend Shared Function GetDEV_BROADCAST_HANDLE(lParam As IntPtr) As DEV_BROADCAST_HANDLE
    '    Return Marshal.PtrToStructure(lParam, GetType(DEV_BROADCAST_VOLUME))
    'End Function

    Friend Shared Function GetDEV_BROADCAST_VOLUME(lParam As IntPtr) As DEV_BROADCAST_VOLUME
        Return Marshal.PtrToStructure(lParam, GetType(DEV_BROADCAST_VOLUME))
    End Function


    ''' <summary>
    ''' Structure with information For RegisterDeviceNotification.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
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

    ''' <summary>
    ''' Struct for parameters of the WM_DEVICECHANGE message
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure DEV_BROADCAST_VOLUME
        Friend dbcv_size As Integer
        Friend dbcv_devicetype As Integer
        Friend dbcv_reserved As Integer
        Friend dbcv_unitmask As Integer
    End Structure

End Class
