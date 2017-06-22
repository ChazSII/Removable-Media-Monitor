Imports System.Text
Imports CS2Soft.RemovableMediaMonitor
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class _DebuggingOnly

    <TestMethod()> Public Sub CatchAll()
        Using devDetector As New DeviceChangeDetector()
            AddHandler devDetector.DeviceArrived, Sub(sender, e)
                                                      If e.Drive = "Q:\" Then
                                                          e.HookQueryRemove = True
                                                      End If

                                                      Debug.WriteLine("New Device: " & e.Drive)
                                                  End Sub
            AddHandler devDetector.DeviceRemoved, Sub(sender, e)
                                                      Debug.WriteLine("Removed Device: " & e.Drive)
                                                  End Sub
            AddHandler devDetector.QueryRemove, Sub(sender, e)
                                                    Debug.WriteLine("Query Device: " & e.Drive)
                                                End Sub


            Dim devNumber As UInt32 = UInt32.MaxValue
            LTR.IO.ImDisk.ImDiskAPI.CreateDevice(0,
                                                 LTR.IO.ImDisk.ImDiskFlags.TypeFile Or
                                                    LTR.IO.ImDisk.ImDiskFlags.Auto Or
                                                    LTR.IO.ImDisk.ImDiskFlags.Removable,
                                                 "imdisktest.img",
                                                 False,
                                                 "Q:",
                                                 devNumber)

            For i As Integer = 0 To 5000 Step 100
                Threading.Thread.Sleep(100)
                System.Windows.Forms.Application.DoEvents()

                Exit For
            Next

            'LTR.IO.ImDisk.ImDiskAPI.RemoveDevice("Q:")

            Do
                System.Windows.Forms.Application.DoEvents()

                Threading.Thread.Sleep(100)
            Loop
        End Using

        Assert.IsTrue(True)
    End Sub

End Class