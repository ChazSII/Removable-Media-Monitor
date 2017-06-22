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

    Friend Shared Function GetDEV_BROADCAST_VOLUME(lParam As IntPtr) As DEV_BROADCAST_VOLUME
        Return Marshal.PtrToStructure(lParam, GetType(DEV_BROADCAST_VOLUME))
    End Function

End Class
