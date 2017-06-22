Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Friend Class WndProcWrapper
    Inherits Form

    ' Win32 constants
    Private Const BROADCAST_QUERY_DENY As Integer = &H424D5144
    Private Const WM_DEVICECHANGE As Integer = &H219

    Friend Event DeviceChanged As DeviceOnChangeEventHandler

    ''' <summary>
    ''' Set up the hidden form. 
    ''' </summary>
    Friend Sub New()
        InitializeComponent()

        Me.Show()
    End Sub

    ''' <summary>
    ''' This function receives all the windows messages for this window (form).
    ''' We call the DriveDetector from here so that is can pick up the messages about
    ''' drives arrived and removed.
    ''' </summary>
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If m.Msg = WM_DEVICECHANGE Then
            ' WM_DEVICECHANGE can have several meanings depending on the WParam value...

            Dim devType As DeviceTypes = Marshal.ReadInt32(m.LParam, 4)
            Dim devEvent As DeviceEvent = m.WParam.ToInt32()

            Dim driveLetter As String = Nothing

            Select Case devEvent
                ' New device has just arrived or has been removed
                Case DeviceEvent.DeviceArrival, DeviceEvent.DeviceRemoveComplete
                    Dim vol As DEV_BROADCAST_VOLUME = NativeMethods.GetDEV_BROADCAST_VOLUME(m.LParam)

                    ' Get the drive letter 
                    driveLetter = [Enum].GetName(GetType(LetterFlags), vol.dbcv_unitmask) & ":\"

                ' Device is about to be removed
                Case DeviceEvent.DeviceQueryRemove
                    If devType <> DeviceTypes.Handle Then
                        Exit Sub
                    End If
            End Select

            Dim eventArgs As New DeviceOnChangeEventArgs With {.DeviceEvent = devEvent, .DeviceType = devType, .Drive = driveLetter}

            RaiseEvent DeviceChanged(Me, eventArgs)

            ' If the client wants to cancel, let Windows know
            If eventArgs.Cancel Then
                m.Result = CType(BROADCAST_QUERY_DENY, IntPtr)
            End If
        End If
    End Sub

#Region "Form Setup"
    Private label1 As Label

    Private Sub WndProcWrapper_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Create really small form, invisible anyway.
        Me.Size = New Drawing.Size(5, 5)
        Me.Visible = False
    End Sub

    Private Sub InitializeComponent()
        Me.label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(13, 30)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(314, 13)
        Me.label1.TabIndex = 0
        Me.label1.Text = "This is invisible form. To see DriveDetector code click View Code"
        '
        'WndProcWrapper
        '
        Me.ClientSize = New System.Drawing.Size(360, 80)
        Me.Controls.Add(Me.label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DetectorForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

#End Region

End Class

