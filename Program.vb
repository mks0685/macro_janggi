Imports System.Windows.Forms

Namespace MacroAutoControl
    Module Program
        <STAThread>
        Sub Main(args As String())
            If args.Length > 0 AndAlso args(0) = "--diag" Then
                GlowDiag.RunDiagnostic()
                Return
            End If
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New MainForm())
        End Sub
    End Module
End Namespace
