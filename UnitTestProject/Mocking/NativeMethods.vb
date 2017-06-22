Imports System.Runtime.InteropServices

Namespace Mocking

    Friend Class NativeMethods

        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function SendMessageTimeout(ByVal hWnd As IntPtr,
                                          ByVal msg As Integer,
                                          ByVal wParam As IntPtr,
                                          ByVal lParam As IntPtr,
                                          ByVal flags As SendMessageTimeoutFlags,
                                          ByVal timeout As Integer,
                                          ByRef result As IntPtr) As IntPtr
        End Function
    End Class

End Namespace