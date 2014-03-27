Imports System.Runtime.InteropServices

Public Class MainForm

    <DllImport("user32")>
    Private Shared Function RegisterHotKey(handle As IntPtr, id As Integer, fsModifier As UInt32, vk As UInt32) As Boolean
    End Function
    <DllImport("user32")>
    Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    Private KeyId As Integer
    Private WithEvents ScreenCaptureForm As ScreenCaptureForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not RegisterHotKey(Me.Handle, KeyId, 0, Keys.PrintScreen) Then
            MessageBox.Show("PrintScreen は他のアプリに関連付けられています")
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_HOTKEY As Integer = &H312
        MyBase.WndProc(m)

        Select Case m.Msg
            Case WM_HOTKEY
                If m.WParam.ToInt32 = KeyId Then
                    ShowScreenCaptureForm()
                End If
            Case Else
        End Select

    End Sub

    Private Sub MainForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        UnregisterHotKey(Me.Handle, KeyId)
    End Sub

    Private Sub ShowScreenCaptureForm()
        If ScreenCaptureForm Is Nothing Then
            ScreenCaptureForm = New ScreenCaptureForm With {.TopMost = True}
            ScreenCaptureForm.Show()
        End If
    End Sub

    Private Sub ScreenCaptureForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ScreenCaptureForm.FormClosed
        If ScreenCaptureForm.DialogResult = Windows.Forms.DialogResult.OK Then
            Label1.Text = ScreenCaptureForm.Rectangle.ToString
        End If

        ScreenCaptureForm.Dispose()
        ScreenCaptureForm = Nothing


    End Sub

End Class
