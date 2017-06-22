Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports CS2Soft.RemovableMediaMonitor
Imports UnitTestProject.Mocking

<TestClass()> Public Class WndProcEvents_UnitTest

    <TestMethod()> Public Sub TestDeviceInsert()
        Dim eventRaised As Boolean = False
        Dim errorRaised As Boolean = False

        Dim device As New MockedDevice(LetterFlags.O)

        Using devDetector As New DeviceChangeDetector()
            AddHandler devDetector.DeviceArrived, Sub(sender, e)
                                                      Dim eventLetter As String = e.Drive(0)
                                                      Dim deviceLetter As String = [Enum].GetName(GetType(LetterFlags),
                                                                                                 device.DriveLetter)
                                                      If eventLetter = deviceLetter Then
                                                          eventRaised = True
                                                      End If
                                                  End Sub

            errorRaised = Not device.Insert(devDetector.Handle)

            For delay As Integer = 0 To 50
                System.Windows.Forms.Application.DoEvents()

                If eventRaised Then Exit For

                Threading.Thread.Sleep(100)
            Next

            Assert.IsTrue(eventRaised)
            Assert.IsFalse(errorRaised)
        End Using
    End Sub

    <TestMethod()> Public Sub TestDeviceQueryRemove()
        Dim eventRaised As Boolean = False
        Dim errorRaised As Boolean = False

        Dim device As New MockedDevice(LetterFlags.O)

        Using devDetector As New DeviceChangeDetector()
            AddHandler devDetector.QueryRemove, Sub(sender, e) eventRaised = True

            errorRaised = Not (device.SafeEject(devDetector.Handle, devDetector.NotifyHandle) = &H1)

            For delay As Integer = 0 To 50
                System.Windows.Forms.Application.DoEvents()

                If eventRaised Then Exit For

                Threading.Thread.Sleep(100)
            Next

            Assert.IsTrue(eventRaised)
            Assert.IsFalse(errorRaised)
        End Using
    End Sub

    <TestMethod()> Public Sub TestDeviceQueryRemoveCancel()
        Dim eventRaised As Boolean = False
        Dim errorRaised As Boolean = False

        Dim device As New MockedDevice(LetterFlags.O)

        Using devDetector As New DeviceChangeDetector()
            AddHandler devDetector.QueryRemove, Sub(sender, e)
                                                    e.Cancel = True
                                                    eventRaised = True
                                                End Sub

            errorRaised = Not (device.SafeEject(devDetector.Handle, devDetector.NotifyHandle) = BROADCAST_QUERY_DENY)

            For delay As Integer = 0 To 50
                System.Windows.Forms.Application.DoEvents()

                If eventRaised Then Exit For

                Threading.Thread.Sleep(100)
            Next

            Assert.IsTrue(eventRaised)
            Assert.IsFalse(errorRaised)
        End Using
    End Sub

    <TestMethod()> Public Sub TestDeviceRemoved()
        Dim eventRaised As Boolean = False
        Dim errorRaised As Boolean = False

        Dim device As New MockedDevice(LetterFlags.O)

        Using devDetector As New DeviceChangeDetector()
            AddHandler devDetector.DeviceRemoved, Sub(sender, e)
                                                      Dim eventLetter As String = e.Drive(0)
                                                      Dim deviceLetter As String = [Enum].GetName(GetType(LetterFlags),
                                                                                                 device.DriveLetter)
                                                      If eventLetter = deviceLetter Then
                                                          eventRaised = True
                                                      End If
                                                  End Sub

            errorRaised = Not device.Removed(devDetector.Handle)

            For delay As Integer = 0 To 50
                System.Windows.Forms.Application.DoEvents()

                If eventRaised Then Exit For

                Threading.Thread.Sleep(100)
            Next

            Assert.IsTrue(eventRaised)
            Assert.IsFalse(errorRaised)
        End Using
    End Sub

End Class