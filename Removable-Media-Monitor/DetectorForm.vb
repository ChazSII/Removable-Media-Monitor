''==================================================================''
''                                                                  ''
''      Copied and modified from Jan Dolinay's Drive Detector       ''
''          - http://www.codeproject.com/Articles/18062/            ''
''                                                                  ''
''  Translated from C# to VB via Snippet Converter and manually     ''
''  - http://codeconverter.sharpdevelop.net/SnippetConverter.aspx   ''
''                                                                  ''
''==================================================================''
Imports System.Windows.Forms
' DriveDetector - rev. 1, Oct. 31 2007

''' <summary>
''' Hidden Form which we use to receive Windows messages about flash drives
''' </summary>
Friend Class DetectorForm
    Inherits Form
    Private label1 As Label
    Private detector As DriveDetector

    ''' <summary>
    ''' Set up the hidden form. 
    ''' </summary>
    ''' <param name="detector">DriveDetector object which will receive notification about USB drives, see WndProc</param>
    Public Sub New(detector As DriveDetector)
        InitializeComponent()

        detector = detector
    End Sub

    Private Sub DetectorForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Create really small form, invisible anyway.
        Me.Size = New Drawing.Size(5, 5)
        Me.Visible = False
    End Sub

    ''' <summary>
    ''' This function receives all the windows messages for this window (form).
    ''' We call the DriveDetector from here so that is can pick up the messages about
    ''' drives arrived and removed.
    ''' </summary>
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If detector IsNot Nothing Then
            detector.WndProc(m)
        End If
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
        'DetectorForm
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

End Class
