Public Class DeviceChangeDetector
    Implements IDisposable


    '' Events signalized to the client app.
    '' Add handlers for these events in your form to be notified of removable device events 
    Public Event DeviceArrived As DeviceOnChangeEventHandler
    Public Event DeviceRemoved As DeviceOnChangeEventHandler
    Public Event QueryRemove As DeviceOnChangeEventHandler

    Private WithEvents WndProc As WndProcWrapper
    Private WithEvents DevNotify As DeviceNotifyWrapper

    ''' <summary>
    ''' The easiest way to use DriveDetector. 
    ''' It will create hidden form for processing Windows messages about USB drives
    ''' You do not need to override WndProc in your form.
    ''' </summary>
    Public Sub New()
        WndProc = New WndProcWrapper()
        DevNotify = New DeviceNotifyWrapper(WndProc.Handle)
    End Sub

    ''' <summary>
    ''' Consructs DriveDetector object setting also path to file which should be opened
    ''' when registering for query remove.  
    ''' </summary>
    ''' <param name="FileToOpen">Optional. Name of a file on the removable drive which should be opened. 
    ''' If null, root directory of the drive will be opened. Opening a file is needed for us 
    ''' to be able to register for the query remove message. TIP: For files use relative path without drive letter.
    ''' e.g. "SomeFolder\file_on_flash.txt"</param>
    Public Sub New(FileToOpen As String)
        WndProc.Show()
        DevNotify = New DeviceNotifyWrapper(WndProc.Handle, FileToOpen)
    End Sub

    Private Sub WndProc_DeviceChanged(sender As Object, e As DeviceOnChangeEventArgs) Handles WndProc.DeviceChanged
        Select Case e.DeviceEvent
                ' New device has just arrived
            Case DeviceEvent.DeviceArrival
                If e.DeviceType = DeviceTypes.Volume Then
                    ' Call the client event handler
                    RaiseEvent DeviceArrived(Me, e)

                    ' Register for query remove if requested
                    If e.HookQueryRemove Then
                        ' If something is already hooked, unhook it now
                        If DevNotify.IsQueryHooked Then
                            DevNotify.DisableQueryRemove()
                        End If

                        DevNotify.EnableQueryRemove(e.Drive)
                    End If
                End If
            Case DeviceEvent.DeviceQueryRemove
                ' Device is about to be removed
                ' Any application can cancel the removal

                If e.DeviceType = DeviceTypes.Handle Then
                    e.Drive = DevNotify.CurrentDevice

                    ' Call the event handler in client
                    RaiseEvent QueryRemove(Me, e)

                    ' If the client wants to cancel, let Windows know
                    If Not e.Cancel Then
                        ' Change 28.10.2007: Unregister the notification, this will
                        ' close the handle to file or root directory also. 
                        ' We have to close it anyway to allow the removal so
                        ' even if some other app cancels the removal we would not know about it...                                    
                        ' will also close the mFileOnFlash
                        DevNotify.DisableQueryRemove()
                    End If
                End If
            Case DeviceEvent.DeviceRemoveComplete
                ' Device has been removed

                If e.DeviceType = DeviceTypes.Volume Then
                    ' Call the client event handler
                    RaiseEvent DeviceRemoved(Me, e)
                End If
        End Select
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                WndProc.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.

            ' Unregister and close the file we may have opened on the removable drive. 
            DevNotify.DisableQueryRemove()

        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#If CONFIG = "Debug" Then
    Public ReadOnly Property Handle As IntPtr
        Get
            Return WndProc.Handle
        End Get
    End Property

    Public ReadOnly Property NotifyHandle As IntPtr
        Get
            Return DevNotify.NotifyHandle
        End Get
    End Property
#End If

End Class
