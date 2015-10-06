Option Explicit On
Option Infer Off
Public Class Form1
    Dim init As Boolean = False
    Dim mediaPlayers(20) As AxWMPLib.AxWindowsMediaPlayer
    Dim labels(20) As Label
    Dim lastAuto As Integer
    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        mediaPlayers(0) = AxWindowsMediaPlayer1
        mediaPlayers(1) = AxWindowsMediaPlayer2
        mediaPlayers(2) = AxWindowsMediaPlayer3
        mediaPlayers(3) = AxWindowsMediaPlayer4
        mediaPlayers(4) = AxWindowsMediaPlayer5
        mediaPlayers(5) = AxWindowsMediaPlayer6
        mediaPlayers(6) = AxWindowsMediaPlayer7
        mediaPlayers(7) = AxWindowsMediaPlayer8
        mediaPlayers(8) = AxWindowsMediaPlayer9
        mediaPlayers(9) = AxWindowsMediaPlayer10
        mediaPlayers(10) = AxWindowsMediaPlayer11
        mediaPlayers(11) = AxWindowsMediaPlayer12
        mediaPlayers(12) = AxWindowsMediaPlayer13
        mediaPlayers(13) = AxWindowsMediaPlayer14
        mediaPlayers(14) = AxWindowsMediaPlayer15
        mediaPlayers(15) = AxWindowsMediaPlayer16
        mediaPlayers(16) = AxWindowsMediaPlayer17
        mediaPlayers(17) = AxWindowsMediaPlayer18
        mediaPlayers(18) = AxWindowsMediaPlayer19
        mediaPlayers(19) = AxWindowsMediaPlayer20
        labels(0) = Label1
        labels(1) = Label2
        labels(2) = Label3
        labels(3) = Label4
        labels(4) = Label5
        labels(5) = Label6
        labels(6) = Label7
        labels(7) = Label8
        labels(8) = Label9
        labels(9) = Label10
        labels(10) = Label11
        labels(11) = Label12
        labels(12) = Label13
        labels(13) = Label14
        labels(14) = Label15
        labels(15) = Label16
        labels(16) = Label17
        labels(17) = Label18
        labels(18) = Label19
        labels(19) = Label20
    End Sub

    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim x As WMPLib.IWMPMedia =
        'wmpLeft.currentMedia = wmpLeft.newMedia("E:\Personnel\Musiques\La soupe aux choux.mp3")
        'wmpLeft.Ctlcontrols.play()
        'wmpRight.currentMedia = wmpMix.newMedia("E:\Personnel\Musiques\Sing Halleluja.mp3")
        'wmpRight.Ctlcontrols.play()

    End Sub

    Private Sub tmrMix_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If wmpLeft.settings.volume > 0 Then wmpLeft.settings.volume -= 1
        'If wmpMix.settings.volume < 100 Then wmpMix.settings.volume += 1
        'Try
        ''    Me.Text = wmpLeft.Ctlcontrols.currentPosition & " / " & wmpLeft.currentMedia.duration & " | " & wmpRight.Ctlcontrols.currentPosition & " / " & wmpRight.currentMedia.duration
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub TableLayoutPanel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TrackBar1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TrackBar1.ValueChanged
        'Try
        Try
            mediaPlayers(ComboBox1.Text).settings.volume = 100 - TrackBar1.Value
        Catch ex As Exception
        End Try

        'Catch ex As System.Windows.Forms.AxHost.InvalidActiveXStateException
        'End Try
        Try
            mediaPlayers(ComboBox2.Text).settings.volume = TrackBar1.Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If init Then Exit Sub
        Dim i As Integer
        For i = 0 To 19
            AddHandler mediaPlayers(i).PlayStateChange, AddressOf _PlayStateChange
            labels(i).AllowDrop = True
            labels(i).Text = "En attente ..."
            AddHandler labels(i).DragOver, AddressOf _DragOver
            AddHandler labels(i).DragEnter, AddressOf _DragEnter
            AddHandler labels(i).DragDrop, AddressOf _DragDrop
            labels(i).Dock = DockStyle.Fill
            mediaPlayers(i).Dock = DockStyle.Fill
            mediaPlayers(i).settings.autoStart = False
            mediaPlayers(i).settings.volume = 0
            mediaPlayers(i).settings.mute = False
        Next
        init = True
        ComboBox1.Text = 0
        ComboBox2.Text = 1
        cmbAuto.Text = 0
        'oThread.Start()
    End Sub

    Private Sub wmpLeft_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub wmpLeft_PositionChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PositionChangeEvent)
    End Sub

    Private Sub AxWindowsMediaPlayer1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxWindowsMediaPlayer1.Enter

    End Sub

    Private Sub _PlayStateChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) 'Handles _
        'AxWindowsMediaPlayer1.PlayStateChange, AxWindowsMediaPlayer2.PlayStateChange, AxWindowsMediaPlayer3.PlayStateChange, AxWindowsMediaPlayer4.PlayStateChange, _
        'AxWindowsMediaPlayer5.PlayStateChange, AxWindowsMediaPlayer6.PlayStateChange, AxWindowsMediaPlayer7.PlayStateChange, AxWindowsMediaPlayer8.PlayStateChange, _
        'AxWindowsMediaPlayer9.PlayStateChange, AxWindowsMediaPlayer10.PlayStateChange, AxWindowsMediaPlayer11.PlayStateChange, AxWindowsMediaPlayer12.PlayStateChange, _
        'AxWindowsMediaPlayer13.PlayStateChange, AxWindowsMediaPlayer14.PlayStateChange, AxWindowsMediaPlayer15.PlayStateChange, AxWindowsMediaPlayer16.PlayStateChange, _
        'AxWindowsMediaPlayer17.PlayStateChange, AxWindowsMediaPlayer18.PlayStateChange, AxWindowsMediaPlayer19.PlayStateChange, AxWindowsMediaPlayer20.PlayStateChange
        'CheckBox1.Checked = False
        'If chkAuto.Checked AndAlso Not (e.newState = WMPLib.WMPPlayState.wmppsPlaying Or e.newState = WMPLib.WMPPlayState.wmppsPaused) Then
        If chkAuto.Checked AndAlso (e.newState = WMPLib.WMPPlayState.wmppsStopped) Then
            If ListBox1.Items.Count > 0 Then
                Dim mp = TryCast(sender, AxWMPLib.AxWindowsMediaPlayer)
                Dim i = Array.IndexOf(mediaPlayers, mp)
                Dim o = ListBox1.Items(0)
                If ListBox1.Items.Count > 0 Then
                    ListBox1.Items.Remove(o)
                    setFileIndex(i, o)
                Else
                    setFileIndex(i, "")
                    'labels(i Mod 20).Text = "En attente ..."
                    'mp.currentMedia = mp.newMedia("")
                End If
            End If
        End If
        'If e.newState = WMPLib.WMPPlayState.wmppsPlaying Then
        '    For i = 0 To 19
        '        Dim mediaPlayer = mediaPlayers(i)
        '        If sender Is mediaPlayer Then
        '        Else
        '            If mediaPlayer.playState = WMPLib.WMPPlayState.wmppsPlaying Then
        '                mediaPlayer.Ctlcontrols.pause()
        '            End If
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub _DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        'e.Effect = DragDropEffects.Copy
        'TryCast(sender, Label).Text = e.Data
        'MsgBox(e.Data.GetType())
    End Sub

    Private Sub _DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Dim d = DirectCast(e.Data, DataObject)
        If d.ContainsFileDropList AndAlso d.GetFileDropList.Count = 1 Then
            'MsgBox(d.GetFileDropList(0))
            'DirectCast(Me.DataContext, ViewModel.OptionsPanelViewModel).TimeController.EndMediaUri = New Uri(d.GetFileDropList(0))
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent("WMP DND EXTERNAL FORMAT") Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent(DataFormats.Text) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
        'If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
        '    e.Effect = DragDropEffects.Copy
        'Else
        '    e.Effect = DragDropEffects.None
        'End If
    End Sub
    Private Sub setFileIndex(ByVal i As Integer, ByVal s As String)
        Dim m As WMPLib.IWMPMedia = mediaPlayers(i Mod 20).newMedia(s)
        mediaPlayers(i Mod 20).currentMedia = m
        labels(i Mod 20).Text = (i Mod 20) & " : " & m.name & vbCrLf & m.sourceURL
    End Sub
    Private Function searchLabel(ByVal o As Object) As Integer
        Dim i As Integer
        For i = 0 To 19
            If labels(i) Is o Then Return freeIndexOrNext(i)
        Next
        Return -1
    End Function

    Private Function nextIndex(ByVal i As Integer) As Integer
        Return freeIndexOrNext(i + 1)
    End Function
    Private Function nextIndexWOF(ByVal i As Integer) As Integer
        Return freeIndexOrNextWOF(i)
    End Function

    Private Function freeIndexOrNext(ByVal i As Integer) As Integer
        If mediaPlayers(i Mod 20).playState = WMPLib.WMPPlayState.wmppsPlaying Then Return freeIndexOrNext(i + 1)
        Return i
    End Function

    Private Function freeIndexOrNextWOF(ByVal i As Integer) As Integer
        If mediaPlayers(i Mod 20).currentMedia Is Nothing Then
        ElseIf mediaPlayers(i Mod 20).currentMedia.sourceURL.Length < 10 Then
        ElseIf i = 40 Then
            Return -1
        Else
            Return freeIndexOrNextWOF(i + 1)
        End If
        Return i
    End Function

    Private Sub _DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Dim d = DirectCast(e.Data, DataObject)
        If d.ContainsFileDropList Then
            setFileIndex(searchLabel(sender), d.GetFileDropList(0))
            'DirectCast(Me.DataContext, ViewModel.OptionsPanelViewModel).TimeController.EndMediaUri = New Uri(d.GetFileDropList(0))
        ElseIf d.GetDataPresent("WMP DND EXTERNAL FORMAT") Then
            Try
                Dim stream As IO.MemoryStream = DirectCast(d.GetData("WMP DND EXTERNAL FORMAT"), IO.MemoryStream)
                Using st As New IO.StreamReader(stream, System.Text.Encoding.Unicode)
                    Dim txt = st.ReadToEnd
                    'Dim x As WMPLib.IWMPPlaylist
                    'MsgBox(DirectCast(d.GetData("WMP DND EXTERNAL FORMAT"), WMPLib.IWMPPlaylist))
                    Dim pos As Integer = 0
                    Dim c = searchLabel(sender)
                    While True
                        pos = InStr(pos + 1, txt, ":\")
                        If pos > 0 Then
                            Dim part = Mid(txt, pos - 1)
                            part = Mid(part, 1, InStr(part, Chr(0)) - 1)
                            setFileIndex(c, part)
                            c += 1
                            'Debug.Print(part)
                            'MsgBox(part.ToCharArray())
                        Else
                            Exit While
                        End If
                    End While
                    'MsgBox(txt.to)
                End Using
            Catch ex As Exception
            End Try
        ElseIf d.GetDataPresent(DataFormats.Text) Then
                setFileIndex(searchLabel(sender), d.GetData(DataFormats.Text))
        End If
    End Sub

    Private Sub ActiveMedia_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged, ComboBox2.TextChanged
        Dim m1 = ComboBox1.Text
        Dim m2 = ComboBox2.Text
        If m1 = "" OrElse m2 = "" Then Exit Sub
        Dim i As Integer
        For i = 0 To 19
            If i = m1 Then
            ElseIf i = CInt(m2) Then
            Else
                mediaPlayers(i).settings.volume = 0
            End If
        Next
        Dim v1 = 0
        Dim v2 = 0
        Try
            v1 = mediaPlayers(m1).settings.volume
        Catch ex As Exception
        End Try
        Try
            v2 = mediaPlayers(m2).settings.volume
        Catch ex As Exception
        End Try
        If sender Is ComboBox1 Then
            If v1 + v2 = 0 Then v2 = 100
            TrackBar1.Value = v2
        Else
            If v1 + v2 = 0 Then v1 = 100
            TrackBar1.Value = (100 - v1)
        End If
        TrackBar1_ValueChanged(Nothing, Nothing)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAuto.CheckedChanged
        'If cmbAuto.SelectedIndex is Nothing then cmbAuto.SelectedIndex = 0
        If chkAuto.Checked Then
            Dim first As Integer = (CInt(cmbAuto.Text)) Mod 20
            Dim j As Integer = first
            Dim i As Integer
            For i = 0 To 19
                j = nextIndexWOF(j)
                If j = -1 Then
                    Exit For
                ElseIf i = 0 Then
                ElseIf j = first Then
                    Exit For
                End If
                If ListBox1.Items.Count > 0 Then
                    Dim o = ListBox1.Items(0)
                    ListBox1.Items.Remove(o)
                    setFileIndex(j, o)
                End If
            Next
            
        End If
        Timer1.Enabled = chkAuto.Checked

    End Sub
    'Dim oThread As New Threading.Thread(AddressOf ThreadProc)
    'Public Sub ThreadProc()

    'End Sub
    Private Sub ensurePlayNext(ByVal i As Integer)
        'i+1
        'If mediaPlayers((i + 1) Mod 20).currentMedia
        Dim j As Integer
        For j = 1 To 18
            If Not mediaPlayers((i + j) Mod 20).currentMedia Is Nothing Then
                If Not mediaPlayers((i + j) Mod 20).playState = WMPLib.WMPPlayState.wmppsPlaying Then
                    mediaPlayers((i + j) Mod 20).Ctlcontrols.play()
                    If ComboBox1.Text = cmbAuto.Text Then
                        ComboBox2.Text = (i + j) Mod 20
                    Else
                        ComboBox1.Text = (i + j) Mod 20
                    End If
                    cmbAuto.Text = (i + j) Mod 20
                End If
                Exit For
            End If
        Next
    End Sub

    Private Sub ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
        Dim d = DirectCast(e.Data, DataObject)
        If d.ContainsFileDropList Then
            Dim x = d.GetFileDropList
            Dim i As Integer
            For i = 0 To x.Count - 1
                If indexOfItemUnderMouseToDrop <> ListBox.NoMatches Then
                    ListBox1.Items.Insert(indexOfItemUnderMouseToDrop, x.Item(i))
                Else
                    ListBox1.Items.Add(x.Item(i))
                End If
            Next
            ''setFileIndex(searchLabel(sender), d.GetFileDropList(0))
            'DirectCast(Me.DataContext, ViewModel.OptionsPanelViewModel).TimeController.EndMediaUri = New Uri(d.GetFileDropList(0))
        ElseIf d.GetDataPresent("WMP DND EXTERNAL FORMAT") Then
            Dim stream As IO.MemoryStream = DirectCast(d.GetData("WMP DND EXTERNAL FORMAT"), IO.MemoryStream)
            Using st As New IO.StreamReader(stream, System.Text.Encoding.Unicode)
                Dim txt = st.ReadToEnd
                'Dim x As WMPLib.IWMPPlaylist
                'MsgBox(DirectCast(d.GetData("WMP DND EXTERNAL FORMAT"), WMPLib.IWMPPlaylist))
                Dim pos As Integer = 0
                'Dim c = searchLabel(sender)
                While True
                    pos = InStr(pos + 1, txt, ":\")
                    If pos > 0 Then
                        Dim part = Mid(txt, pos - 1)
                        part = Mid(part, 1, InStr(part, Chr(0)) - 1)
                        If indexOfItemUnderMouseToDrop <> ListBox.NoMatches Then
                            ListBox1.Items.Insert(indexOfItemUnderMouseToDrop, part)
                        Else
                            ListBox1.Items.Add(part)
                        End If
                        'setFileIndex(c, part)
                        'c += 1
                        'Debug.Print(part)
                        'MsgBox(part.ToCharArray())
                    Else
                        Exit While
                    End If
                End While
                'MsgBox(txt.to)
            End Using
        ElseIf d.GetDataPresent(DataFormats.Text) Then
            If indexOfItemUnderMouseToDrop <> ListBox.NoMatches Then
                ListBox1.Items.Insert(indexOfItemUnderMouseToDrop, d.GetData(DataFormats.Text))
            Else
                ListBox1.Items.Add(d.GetData(DataFormats.Text))
            End If
        End If
        lastDropped = indexOfItemUnderMouseToDrop
        indexOfItemUnderMouseToDrop = ListBox.NoMatches
    End Sub
    Dim lastDropped As Integer = ListBox.NoMatches
    Private Sub ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
        lastDropped = ListBox.NoMatches
        indexOfItemUnderMouseToDrop = ListBox.NoMatches
        Dim d = DirectCast(e.Data, DataObject)
        If d.ContainsFileDropList Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent("WMP DND EXTERNAL FORMAT") Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent(DataFormats.Text) AndAlso (e.AllowedEffect And DragDropEffects.Copy) = DragDropEffects.Copy Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent(DataFormats.Text) AndAlso (e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move Then
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub ListBox1_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DragLeave
        Text = Now
    End Sub
    Dim indexOfItemUnderMouseToDrop As Integer
    Private Sub ListBox1_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragOver
        lastDropped = ListBox.NoMatches
        indexOfItemUnderMouseToDrop = ListBox1.IndexFromPoint(ListBox1.PointToClient(New Point(e.X, e.Y)))
        Dim d = DirectCast(e.Data, DataObject)
        If d.ContainsFileDropList Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent("WMP DND EXTERNAL FORMAT") Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent(DataFormats.Text) AndAlso (e.AllowedEffect And DragDropEffects.Copy) = DragDropEffects.Copy Then
            e.Effect = DragDropEffects.Copy
        ElseIf d.GetDataPresent(DataFormats.Text) AndAlso (e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move Then
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub ListBox1_GiveFeedback(ByVal sender As Object, ByVal e As System.Windows.Forms.GiveFeedbackEventArgs) Handles ListBox1.GiveFeedback

    End Sub
    Dim indexOfItemUnderMouseToDrag As Integer
    Private Sub ListBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDown
        indexOfItemUnderMouseToDrag = ListBox1.IndexFromPoint(e.X, e.Y)

   
    End Sub

    Private Sub ListBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseMove
        If ((e.Button And MouseButtons.Left) = MouseButtons.Left) Then
            Dim Index As Integer
            Dim o
            Index = ListBox1.IndexFromPoint(e.X, e.Y)
            If Index >= 0 Then
                o = ListBox1.Items(Index)
                If o Is Nothing Then Exit Sub
            Else
                Exit Sub
            End If
            lastDropped = ListBox.NoMatches
            Dim v As Integer = ListBox1.DoDragDrop(o, DragDropEffects.Move)
            If v And DragDropEffects.Move = DragDropEffects.Move Then
                If lastDropped <> ListBox.NoMatches Then
                    If lastDropped < indexOfItemUnderMouseToDrag Then
                        indexOfItemUnderMouseToDrag += 1
                    End If
                End If
                ListBox1.Items.RemoveAt(indexOfItemUnderMouseToDrag)
                lastDropped = ListBox.NoMatches
            End If
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem Is Nothing Then Exit Sub
        Dim o = ListBox1.SelectedItem
        Dim i = ListBox1.SelectedIndex
        ListBox1.Items.Remove(o)
        ListBox1.Items.Insert(0, o)
        ListBox1.SelectedIndex = i
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ListBox1.SelectedItem Is Nothing Then Exit Sub
        Dim o = ListBox1.SelectedItem
        Dim i = ListBox1.SelectedIndex
        If i > 0 Then
            ListBox1.Items.Remove(o)
            ListBox1.Items.Insert(i - 1, o)
        End If
        ListBox1.SelectedItem = o
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If ListBox1.SelectedItem Is Nothing Then Exit Sub
        Dim o = ListBox1.SelectedItem
        Dim i = ListBox1.SelectedIndex
        If i < ListBox1.Items.Count - 1 Then
            ListBox1.Items.Remove(o)
            ListBox1.Items.Insert(i + 1, o)
        End If
        ListBox1.SelectedItem = o
    End Sub

    Private Sub cmbAuto_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAuto.SelectedIndexChanged

    End Sub

    Private Sub cmbAuto_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAuto.SelectedValueChanged

    End Sub

    Private Sub cmbAuto_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAuto.TextChanged
        If chkAuto.Checked Then Exit Sub
        Dim oldIndex
        If mediaPlayers(ComboBox1.Text).playState = WMPLib.WMPPlayState.wmppsPlaying Then
            If mediaPlayers(ComboBox2.Text).playState = WMPLib.WMPPlayState.wmppsPlaying Then
                If TrackBar1.Value < 50 Then
                    ' Garder Gauche
                    oldIndex = ComboBox2.Text
                    mediaPlayers(oldIndex).settings.volume = 0
                    mediaPlayers(oldIndex).Ctlcontrols.stop()
                    ComboBox2.Text = cmbAuto.Text
                Else
                    'garder droit
                    oldIndex = ComboBox1.Text
                    mediaPlayers(oldIndex).settings.volume = 0
                    mediaPlayers(oldIndex).Ctlcontrols.stop()
                    ComboBox1.Text = cmbAuto.Text
                End If
            Else
                oldIndex = ComboBox2.Text
                mediaPlayers(oldIndex).settings.volume = 0
                mediaPlayers(oldIndex).Ctlcontrols.stop()
                ComboBox2.Text = cmbAuto.Text
            End If
        Else
            If mediaPlayers(ComboBox2.Text).playState = WMPLib.WMPPlayState.wmppsPlaying Then
                oldIndex = ComboBox1.Text
                mediaPlayers(oldIndex).settings.volume = 0
                mediaPlayers(oldIndex).Ctlcontrols.stop()
                ComboBox1.Text = cmbAuto.Text
            Else
                If ComboBox2.Text = cmbAuto.Text Then
                Else
                    ComboBox1.Text = cmbAuto.Text
                End If
            End If
        End If
        TrackBar1_ValueChanged(Nothing, Nothing)

    End Sub
    Dim action As New Object
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '       While True
        '            While chkAuto.Checked
        'Dim m1 = ComboBox1.SelectedIndex
        'Dim m2 = ComboBox2.SelectedIndex
        'If mediaPlayers(m1).Ctlcontrols.currentPosition > mediaPlayers(m1).currentMedia.duration - 10 Then
        'ElseIf mediaPlayers(m2).Ctlcontrols.currentPosition > mediaPlayers(m2).currentMedia.duration - 10 Then
        'End If
        'If action Then Exit Sub
        'action = True
        SyncLock action
            Try
                If mediaPlayers(cmbAuto.Text).Ctlcontrols.currentPosition < 10 Then
                    If cmbAuto.Text = ComboBox1.Text Then
                        TrackBar1.Value = 100 - (mediaPlayers(cmbAuto.Text).Ctlcontrols.currentPosition * 10)
                    Else
                        TrackBar1.Value = (mediaPlayers(cmbAuto.Text).Ctlcontrols.currentPosition * 10)
                    End If
                ElseIf mediaPlayers(cmbAuto.Text).Ctlcontrols.currentPosition > mediaPlayers(cmbAuto.Text).currentMedia.duration - 10 Then
                    ensurePlayNext(cmbAuto.Text)
                    'TrackBar1.Value = 100 - ((mediaPlayers(cmbAuto.SelectedIndex).currentMedia.duration - mediaPlayers(cmbAuto.SelectedIndex).Ctlcontrols.currentPosition) * 10)
                End If
            Catch ex As Exception
            End Try
        End SyncLock
        'action = False
        'Threading.Thread.Sleep(50)
        'End While
        'Threading.Thread.Sleep(500)
        'End While
    End Sub

    Private Sub FlowLayoutPanel2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles FlowLayoutPanel2.DragDrop

    End Sub

    Private Sub FlowLayoutPanel2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles FlowLayoutPanel2.DragEnter
        e.Effect = DragDropEffects.All
    End Sub

    Private Sub FlowLayoutPanel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles FlowLayoutPanel2.Paint

    End Sub
End Class
