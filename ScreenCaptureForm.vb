Public Class ScreenCaptureForm

    Private CurrentContext As BufferedGraphicsContext
    Private MyBuffer As BufferedGraphics

    Private ClickedPosition As Point
    Private PreviousPosition As Point

    Private ScreenBitmap As Bitmap
    Private GrayScreenBitmap As Bitmap

    Private Dragging As Boolean

    Private _Rectangle As Rectangle
    ReadOnly Property Rectangle As Rectangle
        Get
            Return _Rectangle
        End Get
    End Property



    Private Sub ScreenCaptureForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Location = Screen.PrimaryScreen.Bounds.Location
        Me.Size = Screen.PrimaryScreen.Bounds.Size

        CurrentContext = BufferedGraphicsManager.Current
        MyBuffer = CurrentContext.Allocate(Me.CreateGraphics, Me.DisplayRectangle)

        Dim g As Graphics
        ScreenBitmap = New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        g = Graphics.FromImage(ScreenBitmap)
        g.CopyFromScreen(New Point(0, 0), New Point(0, 0), ScreenBitmap.Size)
        g.Dispose()

        GrayScreenBitmap = ScreenBitmap.Clone
        g = Graphics.FromImage(GrayScreenBitmap)
        g.FillRectangle(New SolidBrush(Color.FromArgb(128, Color.Black)), New Rectangle(New Point(0, 0), GrayScreenBitmap.Size))
        g.Dispose()

        Dim r = New Rectangle(0, 0, GrayScreenBitmap.Width, GrayScreenBitmap.Height)
        MyBuffer.Graphics.DrawImage(GrayScreenBitmap, r, r, GraphicsUnit.Pixel)
    End Sub

    Private Sub ScreenCaptureForm_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        MyBuffer.Render()
    End Sub

    Private Sub ScreenCaptureForm_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If MyBuffer IsNot Nothing Then
            MyBuffer.Dispose()
            MyBuffer = CurrentContext.Allocate(Me.CreateGraphics, Me.DisplayRectangle)
        End If
    End Sub

    Private Sub ScreenCaptureForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        MyBuffer.Dispose()
    End Sub

    Private Sub ScreenCaptureForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ScreenCaptureForm_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        Dragging = True
        ClickedPosition = e.Location
    End Sub

    Private Sub ScreenCaptureForm_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Dim r As Rectangle

        r = New Rectangle(0, 0, ScreenBitmap.Width, ScreenBitmap.Height)
        MyBuffer.Graphics.DrawImage(GrayScreenBitmap, r, r, GraphicsUnit.Pixel)

        MyBuffer.Graphics.DrawLine(Pens.Orange, New Point(0, e.Y), New Point(ScreenBitmap.Width - 1, e.Y))
        MyBuffer.Graphics.DrawLine(Pens.Orange, New Point(e.X, 0), New Point(e.X, ScreenBitmap.Height - 1))

        If Not Dragging Then
            MyBuffer.Render()
            Exit Sub
        End If

        r = CreateRectangle(ClickedPosition, e.Location)
        MyBuffer.Graphics.DrawImage(ScreenBitmap, r, r, GraphicsUnit.Pixel)

        Dim range = 5
        Dim zoom = 10
        Dim srcRect = New RectangleF(e.X - range - 0.5, e.Y - range - 0.5, range * 2 + 1, range * 2 + 1)
        Dim dstRect = New Rectangle(e.X - (range * 2 + 1) * zoom - 10, e.Y + 10, (range * 2 + 1) * zoom, (range * 2 + 1) * zoom)

        MyBuffer.Graphics.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
        MyBuffer.Graphics.CompositingQuality = Drawing2D.CompositingQuality.AssumeLinear
        MyBuffer.Graphics.SmoothingMode = Drawing2D.SmoothingMode.None

        MyBuffer.Graphics.DrawImage(ScreenBitmap, dstRect, srcRect, GraphicsUnit.Pixel)
        MyBuffer.Graphics.DrawRectangle(New Pen(Brushes.Orange), dstRect)

        MyBuffer.Render()
    End Sub

    Private Sub ScreenCaptureForm_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        Dragging = False
        _Rectangle = CreateRectangle(ClickedPosition, e.Location)
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Function CreateRectangle(p1 As Point, p2 As Point)
        Dim r As New Rectangle
        If p2.Y < p1.Y Then
            If p2.X < p1.X Then
                ' クリック位置から左上
                r = New Rectangle(p2, New Size(p1.X - p2.X + 1, p1.Y - p2.Y + 1))
            Else
                ' クリック位置から右上
                r = New Rectangle(New Point(p1.X, p2.Y), New Size(p2.X - p1.X + 1, p1.Y - p2.Y + 1))
            End If
        Else
            If p2.X < p1.X Then
                ' クリック位置から左下
                r = New Rectangle(New Point(p2.X, p1.Y), New Size(p1.X - p2.X + 1, p2.Y - p1.Y + 1))
            Else
                ' クリック位置から右下
                r = New Rectangle(p1, New Size(p2.X - p1.X + 1, p2.Y - p1.Y + 1))
            End If
        End If

        Return r
    End Function

End Class