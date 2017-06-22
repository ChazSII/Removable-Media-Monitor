Imports System.Runtime.InteropServices

Friend Module NativeStructs

    <Flags()>
    Friend Enum LetterFlags
        A = 1
        B = 2
        C = 4
        D = 8
        E = 16
        F = 32
        G = 64
        H = 128
        I = 256
        J = 512
        K = 1024
        L = 2048
        M = 4096
        N = 8192
        O = 16384
        P = 32768
        Q = 65536
        R = 131072
        S = 262144
        T = 524288
        U = 1048576
        V = 2097152
        W = 4194304
        X = 8388608
        Y = 16777216
        Z = 33554432
    End Enum

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
    End Structure
End Module
