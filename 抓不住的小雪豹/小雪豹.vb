Public Class 小雪豹

    Private Sub Dodge(ByVal UserObject As Object)
        Dim MyRand As New System.Random()  '声明随机数类
        Dim PositionX As Integer, PositionY As Integer, LengthX As Integer, LengthY As Integer
        Dim StartX As Integer, StartY As Integer, EachX As Integer, EachY As Integer

        Dim Index As Integer, Counts As Integer

        '确定移动范围和目标位置
        PositionX = MyRand.Next(My.Computer.Screen.Bounds.Width - UserObject.Width)
        PositionY = MyRand.Next(My.Computer.Screen.Bounds.Height - UserObject.Height)

        '计算移动距离
        LengthX = PositionX - UserObject.Left
        LengthY = PositionY - UserObject.Top

        '标记开始位
        StartX = UserObject.Left
        StartY = UserObject.Top

        '计算纵横比例、校正小数、确定方向
        If Math.Abs(LengthX) >= Math.Abs(LengthY) Then
            If LengthY = 0 Then EachX = LengthX \ 1 Else EachX = LengthX \ LengthY
            LengthX = EachX * LengthY
            'PositionX = LengthX + UserObject.Left
            EachX = IIf(LengthX >= 0, Math.Abs(EachX), -Math.Abs(EachX))
            EachY = IIf(LengthY >= 0, 1, -1)
            Counts = Math.Abs(LengthY)
        Else
            If LengthX = 0 Then EachY = LengthY \ 1 Else EachY = LengthY \ LengthX
            LengthY = EachY * LengthX
            'PositionY = LengthY + UserObject.top
            EachY = IIf(LengthY >= 0, Math.Abs(EachY), -Math.Abs(EachY))
            EachX = IIf(LengthX >= 0, 1, -1)
            Counts = Math.Abs(LengthX)
        End If

        '开始移动
        For Index = 0 To Counts
            'Do not run too fast,honey.
            If Index Mod 2 = 0 Then Threading.Thread.Sleep(1)

            UserObject.Left = StartX + CLng(Index * EachX)
            UserObject.Top = StartY + CLng(Index * EachY)
        Next
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            If Not DesignMode Then
                Dim cp As CreateParams = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or Win32.WS_EX_LAYERED
                Return cp
            Else
                Return MyBase.CreateParams
            End If
        End Get
    End Property

    Private Sub 小雪豹_Move(sender As Object, e As EventArgs) Handles Me.MouseEnter, Me.MouseHover
        Dodge(Me)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.TopMost = True
        Static PicIndex As Integer
        PicIndex = IIf(PicIndex = 14, 0, PicIndex + 1)

        Dim BitmapRectangle As New Rectangle(PicIndex * 124, 0, 124, 128)
        Dim NewBitmap As Bitmap = My.Resources.小雪豹Resource.雪豹.Clone(BitmapRectangle, Imaging.PixelFormat.Format32bppArgb)
        SetAlphaPicture(Me, NewBitmap)

        '用Bitmap作为窗体Icon
        Try
            Me.Icon = Icon.FromHandle(NewBitmap.GetHicon)
            Me.ShowIcon = True
        Catch ex As Exception

        End Try
    End Sub

End Class
