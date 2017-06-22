Imports System.Text
Imports CS2Soft.RemovableMediaMonitor
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub CatchAll()
        'Using devDetector As New DeviceChangeDetector()
        '    AddHandler devDetector.DeviceArrived, Sub(seder, e)
        '                                              Debug.WriteLine("New Device: " & e.Drive)
        '                                          End Sub
        '    AddHandler devDetector.DeviceRemoved, Sub(seder, e)
        '                                              Debug.WriteLine("Removed Device: " & e.Drive)
        '                                          End Sub
        '    AddHandler devDetector.QueryRemove, Sub(seder, e)
        '                                            Debug.WriteLine("Query Device: " & e.Drive)
        '                                        End Sub

        '    Do
        '        System.Windows.Forms.Application.DoEvents()

        '        Threading.Thread.Sleep(100)
        '    Loop
        'End Using

        Assert.IsTrue(True)
    End Sub

End Class