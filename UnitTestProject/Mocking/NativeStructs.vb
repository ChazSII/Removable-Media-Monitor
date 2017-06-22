Imports System.Runtime.InteropServices

Namespace Mocking

    Public Module NativeStructs

        Public Const BROADCAST_QUERY_DENY As Integer = &H424D5144

        <Flags()>
        Public Enum LetterFlags
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

        <Flags()>
        Public Enum SendMessageTimeoutFlags
            SMTO_NORMAL = 0
            SMTO_BLOCK = 1
            SMTO_ABORTIFHUNG = 2
            SMTO_NOTIMEOUTIFNOTHUNG = 8
        End Enum

        ''' <summary>
        ''' Struct for parameters of the WM_DEVICECHANGE message
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure DEV_BROADCAST_VOLUME
            Public dbcv_size As Integer
            Public dbcv_devicetype As Integer
            Public dbcv_reserved As Integer
            Public dbcv_unitmask As Integer
        End Structure

        ''' <summary>
        ''' Structure with information For RegisterDeviceNotification.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure DEV_BROADCAST_HANDLE
            Public dbch_size As Integer
            Public dbch_devicetype As Integer
            Public dbch_reserved As Integer
            Public dbch_handle As IntPtr
            Public dbch_hdevnotify As IntPtr
            Public dbch_eventguid As Guid
            Public dbch_nameoffset As Long
            'public byte[] dbch_data[1]; // = new byte[1];
            Public dbch_data As Byte
        End Structure

    End Module

End Namespace