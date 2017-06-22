' Delegate for event handler to handle the device events 
Public Delegate Sub DeviceOnChangeEventHandler(sender As Object, e As DeviceOnChangeEventArgs)

''' <summary>
''' Our class for passing in custom arguments to our event handlers 
''' </summary>
Public Class DeviceOnChangeEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' </summary>
    Public Property DeviceType As DeviceTypes

    ''' <summary>
    ''' </summary>
    Public Property DeviceEvent As DeviceEvent

    ''' <summary>
    ''' Get/Set the value indicating that the event should be cancelled 
    ''' Only in QueryRemove handler.
    ''' </summary>
    Public Property Cancel As Boolean = False

    ''' <summary>
    ''' Drive letter for the device which caused this event 
    ''' </summary>
    Public Property Drive As String = ""

    ''' <summary>
    ''' Set to true in your DeviceArrived event handler if you wish to receive the 
    ''' QueryRemove event for this drive. 
    ''' </summary>
    Public Property HookQueryRemove As Boolean = False
End Class