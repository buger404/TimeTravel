VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Origin As New TimeSpace
    Dim SX As Long, sy As Long
    Dim MoveMode As Integer, Target As TimeSpace, TX As Long, TY As Long, TDX As Long, TDY As Long
    Dim Score As Long
    Private Type Line
        X As Long
        y As Long
        X2 As Long
        Y2 As Long
        time As Long
    End Type
    Dim Lines() As Line
    Dim EneTime As Long
'======================================
Public Sub CopySpace(S As TimeSpace, T As TimeSpace)
    Dim X As Long, y As Long
    For X = 0 To 5
        For y = 0 To 5
            With T.Area(X, y)
                .IsFriend = S.Area(X, y).IsFriend
                .Strength = S.Area(X, y).Strength
                .time = S.Area(X, y).time
            End With
        Next
    Next
End Sub
Public Function CreateSpace(T As TimeSpace) As TimeSpace
    Dim Created As TimeSpace
    If T.Right Is Nothing And Created Is Nothing Then Set T.Right = New TimeSpace: CopySpace T, T.Right: Set Created = T.Right
    If T.Up Is Nothing And Created Is Nothing Then Set T.Up = New TimeSpace: CopySpace T, T.Up: Set Created = T.Up
    If T.Down Is Nothing And Created Is Nothing Then Set T.Down = New TimeSpace: CopySpace T, T.Down: Set Created = T.Down
    If Not Created Is Nothing Then Created.Distance = T.Distance + 1
    Set CreateSpace = Created
End Function
Public Sub DrawTimeSpace(DX As Long, DY As Long, T As TimeSpace)
    Dim X As Long, y As Long, f As Fighter, F2 As Fighter, Colored As Boolean

    If DX >= -180 And DY >= -180 And DX <= GW And DY <= GH Then
        Page.Paint 0, DX, DY, 181, 181, IIf(T.Failed, argb(255, 255, 0, 0), argb(255, 180, 180, 180)), size:=2, style:=1
        If T.Failed Then
            Page.Paint 0, DX, DY, 181, 181, argb(80, 255, 0, 0)
            Page.PaintLine DX, DY, DX + 180, DY + 180, argb(255, 255, 0, 0), 2
            Page.PaintLine DX + 180, DY, DX, DY + 180, argb(255, 255, 0, 0), 2
        End If
        Page.Writes "时空" & T.id, DX, DY + 191, 18, argb(255, 0, 0, 0), 180, 30, align:=StringAlignmentCenter
    End If
    
    If T.Failed Then GoTo NextT
    
    Dim m As Integer, Buff As Long
    
    For X = 0 To 5
        For y = 0 To 5
            If DX >= -30 And DY >= -30 And DX <= GW And DY <= GH Then
                Set f = T.Area(X, y)
                With f
                    Page.Paint 0, DX, DY, 31, 31, IIf(.Strength = 0, IIf(Colored, argb(255, 242, 242, 242), argb(255, 255, 255, 255)), IIf(.IsFriend, argb(255, 0, 176, 240), argb(255, 255, 83, 57)))
                    m = CheckMouse2
                    Buff = 0
                    If .IsFriend And .Strength Then
                        If MoveMode <> 0 And Target Is T And X = TX And y = TY Then
                            Page.Paint 0, DX, DY, 31, 31, argb(128, 0, 0, 0)
                            TDX = DX + 15: TDY = DY + 15
                            If m = mMouseUp Then
                                MoveMode = 0: Set Target = Nothing: m = 0
                            End If
                        End If
                        If MoveMode = 0 Then
                            If m = mMouseUp Then
                                MoveMode = IIf(Mouse.button = 1, 1, 2)
                                Set Target = T: TX = X: TY = y
                            End If
                        End If
                    End If
                    If Not Target Is Nothing Then
                        If Not (Target Is T And (Not T.Down Is Nothing) And (Not T.Right Is Nothing) And (Not T.Up Is Nothing)) Then
                            If MoveMode <> 0 And m > 0 And (Not Target.Area(TX, TY) Is f) Then
                                Page.Paint 0, DX, DY, 31, 31, argb(255, 255, 170, 49)
                                Buff = Target.Area(TX, TY).Strength * IIf(MoveMode = 1, 1, 0.5)
                                Dim RealBuff As Long
                                RealBuff = Buff
                                Buff = Buff - Buff * 0.3 * Abs(Target.Distance - T.Distance)
                                If m = mMouseUp Then
                                    Dim TS As TimeSpace
                                    Set TS = T
                                    If T Is Target Then Set TS = CreateSpace(T)
                                    TS.Area(X, y).Strength = TS.Area(X, y).Strength * IIf(TS.Area(X, y).IsFriend, 1, -1) + Buff
                                    TS.Area(X, y).IsFriend = (TS.Area(X, y).Strength > 0)
                                    TS.Area(X, y).Strength = Abs(TS.Area(X, y).Strength)
                                    If T Is Target Then
                                        TS.Area(TX, TY).Strength = TS.Area(TX, TY).Strength - RealBuff
                                    Else
                                        ReDim Preserve Lines(UBound(Lines) + 1)
                                        With Lines(UBound(Lines))
                                            .X = DX + 15 - SX
                                            .y = DY + 15 - sy
                                            .X2 = TDX - SX
                                            .Y2 = TDY - sy
                                            .time = GetTickCount
                                        End With
                                        Target.Area(TX, TY).Strength = Target.Area(TX, TY).Strength - RealBuff
                                    End If
                                    If TS.Area(X, y).IsFriend Then Score = Score + TS.Area(X, y).Strength
                                    MoveMode = 0
                                End If
                            End If
                        End If
                    End If
                    If .Strength Or Buff Then Page.Writes .Strength * IIf(MoveMode > 0 And Buff > 0, IIf(.IsFriend, 1, -1), 1) + Buff, DX, DY + 5, 10, argb(255, 255, 255, 255), 30, 20, StringAlignmentCenter, FontStyleBold
                    If .IsFriend And .Strength And GetTickCount - .time >= 1000 Then .Strength = .Strength + Int(Rnd * 10): .time = GetTickCount
                    If (Not .IsFriend) And .Strength And GetTickCount - .time >= 3000 Then
                        .time = GetTickCount
                        If X - 1 < 0 Then
                            Score = Score - .Strength
                            T.Failed = True
                            GoTo NextT
                        End If
                        Set F2 = T.Area(X - 1, y)
                        F2.Strength = F2.Strength - .Strength
                        .Strength = 0
                        F2.IsFriend = (F2.Strength > 0)
                        F2.Strength = Abs(F2.Strength)
                        F2.time = GetTickCount
                        If Not Target Is Nothing Then
                            If F2 Is Target.Area(TX, TY) And F2.IsFriend = False Then MoveMode = 0
                        End If
                    End If
                End With
            End If
            Colored = Not Colored
            DY = DY + 30
            If X = 5 And GetTickCount - EneTime >= IIf(Score < 0, 10000, 10000 * Cubic(IIf(Score < 100000, Score / 100000, 1), 1, 1, 1, 0.1)) And Int(Rnd * 100) = 66 Then
                Set f = T.Area(5, Int(Rnd * 6))
                f.Strength = f.Strength * IIf(f.IsFriend, 1, -1) - Int(Rnd * TotalId * 100)
                f.IsFriend = (f.Strength > 0)
                f.Strength = Abs(f.Strength)
                f.time = GetTickCount
                EneTime = GetTickCount
            End If
        Next
        Colored = Not Colored
        DY = DY - 6 * 30
        DX = DX + 30
    Next
    DX = DX - 6 * 30
    
    If (Target Is T And (Not T.Down Is Nothing) And (Not T.Right Is Nothing) And (Not T.Up Is Nothing)) Then
        Page.Paint 0, DX, DY, 181, 181, argb(120, 0, 0, 0)
    End If
    
NextT:
    If Not T.Right Is Nothing Then DrawTimeSpace DX + 220, DY, T.Right
    If Not T.Up Is Nothing Then DrawTimeSpace DX, DY - 250, T.Up
    If Not T.Down Is Nothing Then DrawTimeSpace DX, DY + 250, T.Down
End Sub
Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear '清空画布

    If MoveMode = 1 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(80, 0, 176, 240)
    If MoveMode = 2 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(80, 26, 219, 206)
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim X As Long, y As Long
    X = SX: y = sy
    
    DrawTimeSpace X, y, Origin
    
    For I = 1 To UBound(Lines)
        If I > UBound(Lines) Then Exit Sub
        With Lines(I)
            Page.PaintLine .X + SX, .y + sy, .X2 + SX, .Y2 + sy, argb(120, 255, 0, 0), 2
        End With
        If GetTickCount - Lines(I).time >= 3000 Then Lines(I) = Lines(UBound(Lines)): ReDim Preserve Lines(UBound(Lines) - 1): I = I - 1
    Next
    
    If IsKeyPress(VK_LEFT) Then SX = SX + 10
    If IsKeyPress(VK_RIGHT) Then SX = SX - 10
    If IsKeyPress(VK_UP) Then sy = sy + 10
    If IsKeyPress(VK_DOWN) Then sy = sy - 10
    
    Page.Writes "得分 " & Score, GW / 2, 40, 20, argb(255, 0, 0, 0), align:=StringAlignmentCenter
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets"  '导入资源
    ECore.Add Page, "AppPage"  '创建页面
    
    SX = GW / 2 - 6 * 30 / 2
    sy = GH / 2 - 6 * 30 / 2

    Origin.Area(0, 3).IsFriend = True
    Origin.Area(0, 3).Strength = 100
    
    EneTime = GetTickCount
    
    ReDim Lines(0)
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
