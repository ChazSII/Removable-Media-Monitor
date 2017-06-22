Imports CS2Soft.RemovableMediaMonitor

Namespace Mocking

    Public Class MockedDevice
        Const WM_DEVICECHANGE As Integer = &H219

        Const DeviceArrival = &H8000
        Const DeviceQueryRemove = &H8001
        Const DeviceRemoveComplete = &H8004

        Public ReadOnly Property DriveLetter As LetterFlags

        Public Sub New(driveLetter As LetterFlags)
            _DriveLetter = driveLetter
        End Sub

        Public Function Insert(destHandle As IntPtr) As Boolean
            Dim result As IntPtr = Wnd_proc_call_dev_brdcst_vol(destHandle, DeviceArrival, DeviceTypes.Volume)
            Dim lastErr = Runtime.InteropServices.Marshal.GetLastWin32Error()

            Return CBool(result) And lastErr = 0
        End Function

        Public Function SafeEject(destHandle As IntPtr, notifyHandle As IntPtr) As IntPtr
            Dim result As IntPtr = Wnd_proc_call_dev_brdcst_handle(destHandle, DeviceQueryRemove, notifyHandle)

            Return result
        End Function

        Public Function Removed(destHandle As IntPtr) As Boolean
            Dim result As IntPtr = Wnd_proc_call_dev_brdcst_vol(destHandle, DeviceRemoveComplete, DeviceTypes.Volume)
            Dim lastErr = Runtime.InteropServices.Marshal.GetLastWin32Error()

            Return CBool(result) And lastErr = 0
        End Function


        Private Function Wnd_proc_call_dev_brdcst_vol(handle As IntPtr, eventType As Integer, deviceType As DeviceTypes) As IntPtr
            '' WindowProc(HWND   hwnd,     // handle to window
            ''            UINT   uMsg,     // WM_DEVICECHANGE
            ''            WPARAM wParam,   // device - change event
            ''            LPARAM lParam ); // event-specific data
            Dim dev_brdcst_vol As New DEV_BROADCAST_VOLUME With {
                .dbcv_devicetype = deviceType,
                .dbcv_unitmask = DriveLetter,
                .dbcv_reserved = 0}

            Dim size As Integer = Runtime.InteropServices.Marshal.SizeOf(dev_brdcst_vol)
            dev_brdcst_vol.dbcv_size = size

            Dim buffer As IntPtr = Runtime.InteropServices.Marshal.AllocHGlobal(size)
            Runtime.InteropServices.Marshal.StructureToPtr(dev_brdcst_vol, buffer, False)

            Dim result As IntPtr = IntPtr.Zero

            NativeMethods.SendMessageTimeout(handle,
                                             WM_DEVICECHANGE,
                                             eventType,
                                             buffer,
                                             SendMessageTimeoutFlags.SMTO_ABORTIFHUNG,
                                             5000,
                                             result)

            Return result
        End Function

        Private Function Wnd_proc_call_dev_brdcst_handle(handle As IntPtr, eventType As Integer, notifyhandle As IntPtr) As IntPtr
            '' WindowProc(HWND   hwnd,     // handle to window
            ''            UINT   uMsg,     // WM_DEVICECHANGE
            ''            WPARAM wParam,   // device - change event
            ''            LPARAM lParam ); // event-specific data
            Dim dev_brdcst_handle As New DEV_BROADCAST_HANDLE With {
                .dbch_devicetype = DeviceTypes.Handle,
                .dbch_hdevnotify = notifyhandle,
                .dbch_handle = IntPtr.Zero}

            Dim size As Integer = Runtime.InteropServices.Marshal.SizeOf(dev_brdcst_handle)
            dev_brdcst_handle.dbch_size = size

            Dim buffer As IntPtr = Runtime.InteropServices.Marshal.AllocHGlobal(size)
            Runtime.InteropServices.Marshal.StructureToPtr(dev_brdcst_handle, buffer, False)

            Dim result As IntPtr = IntPtr.Zero

            Dim result2 = NativeMethods.SendMessageTimeout(handle,
                                             WM_DEVICECHANGE,
                                             eventType,
                                             buffer,
                                             SendMessageTimeoutFlags.SMTO_ABORTIFHUNG,
                                             5000,
                                             result)

            Return result
        End Function

    End Class

End Namespace
