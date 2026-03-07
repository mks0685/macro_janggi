Imports System.Threading
Imports System.Windows.Forms

Namespace MacroAutoControl
    Module Program
        <STAThread>
        Sub Main(args As String())
            If args.Length > 0 AndAlso args(0) = "--diag" Then
                GlowDiag.RunDiagnostic()
                Return
            End If

            Dim mtx As New Mutex(True, "MacroAutoControl_SingleInstance", createdNew:=False)
            Dim created As Boolean = mtx.WaitOne(0, False)
            If Not created Then
                MessageBox.Show("이미 실행 중입니다.", "매크로 장기", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                Application.EnableVisualStyles()
                Application.SetCompatibleTextRenderingDefault(False)
                Application.Run(New MainForm())
            Finally
                mtx.ReleaseMutex()
                mtx.Dispose()
            End Try
        End Sub
    End Module
End Namespace
