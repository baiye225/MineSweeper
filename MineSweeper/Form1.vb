Public Class Form1
    Private MineMap(,), TempMineMap(,), IslandNum, IslandPointNum, MineNum,
        MineNumLeft, MineNumLeftDisplay, MineMapHeight, MineMapWidth As Integer
    Private FirstClick As Boolean = False
    Private SingleIslandMap(,) As Integer
    Private IslandMap As New Dictionary(Of Integer, Integer(,))
    Private ImageMine As New Bitmap("mine.png")
    Private ImageHighlight As New Bitmap("mine highlighted.png")
    Private ImageFlag As New Bitmap("flag.png")
    Private ImageNormal As New Bitmap("face normal.jpeg")
    Private ImageSuccess As New Bitmap("face success.jpeg")
    Private ImageFailure As New Bitmap("face failure.jpeg")

    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        'initialize everything
        Form1_Load(sender, e)
    End Sub

    Private Sub Button_Exit_Click(sender As Object, e As EventArgs) Handles Button_Exit.Click
        'exit
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'load mine image
        ImageMine = New Bitmap(ImageMine, Button1.Width, Button1.Height)
        ImageHighlight = New Bitmap(ImageHighlight, Button1.Width, Button1.Height)
        ImageFlag = New Bitmap(ImageFlag, Button1.Width, Button1.Height)

        'initialize mine map
        MineNum = 15
        MineMapHeight = 12
        MineMapWidth = 8
        MineNumLeft = MineNum
        MineNumLeftDisplay = MineNum
        ReDim MineMap(MineMapWidth - 1, MineMapHeight - 1)

        'initialize all buttons
        InitAllButton()

        'generate the whole mine distribution
        'Init_Mine_Map()

        'initialize face
        PictureBox1.Image = ImageNormal
        PictureBox1.SizeMode = 1

        'initialize first mine click marker
        FirstClick = False

        'initialize mine recorder
        Label2.Text = CStr(MineNum) & "/" & CStr(MineNum)

        'distribute mine in the mine map
        'ShowAllMine()

        'distribute mine marker number in the mine map
        'ShowAllMineMarker()
    End Sub

    'generate the whole distribution
    Private Sub Init_Mine_Map(ByVal X0 As Integer, Y0 As Integer)
        ' generate mine
        Dim MineX, MineY As Integer ' position of the current mine
        Dim i, j As Integer
        Dim MineNumeAround As Integer
        Dim MineNumNow As Integer = MineNum

        'initialize mine in the mine map
        While MineNumNow > 0
            'generate random location of a new mine
            MineX = CInt(Rnd() * (MineMapWidth - 1))
            MineY = CInt(Rnd() * (MineMapHeight - 1))

            'check if the mine exists?
            ' -1 -> mine
            ' 0 -> empty
            ' positive number -> the number of mine around the place
            If MineMap(MineX, MineY) <> -1 And Not (MineX = X0 And MineY = Y0) Then
                MineMap(MineX, MineY) = -1
                MineNumNow -= 1
            End If
        End While

        'initialize mine marker number in the mine map
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1

                If MineMap(i, j) <> -1 Then
                    'reset MineNumeAround
                    MineNumeAround = 0

                    'check up
                    If j > 0 Then
                        If MineMap(i, j - 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check up left
                    If j > 0 And i > 0 Then
                        If MineMap(i - 1, j - 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check up right
                    If j > 0 And i < MineMapWidth - 1 Then
                        If MineMap(i + 1, j - 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check down
                    If j < MineMapHeight - 1 Then
                        If MineMap(i, j + 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check down left
                    If j < MineMapHeight - 1 And i > 0 Then
                        If MineMap(i - 1, j + 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check down right
                    If j < MineMapHeight - 1 And i < MineMapWidth - 1 Then
                        If MineMap(i + 1, j + 1) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check left
                    If i > 0 Then
                        If MineMap(i - 1, j) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'check right
                    If i < MineMapWidth - 1 Then
                        If MineMap(i + 1, j) = -1 Then
                            MineNumeAround += 1
                        End If
                    End If

                    'locate current button and setup number
                    If MineNumeAround > 0 Then
                        MineMap(i, j) = MineNumeAround
                    End If
                End If
            Next
        Next

        ' get island map
        GetEmptyIsland()
    End Sub

    ' distribute mine
    Private Sub ShowAllMine()
        Dim i, j As Integer
        Dim MyButton As Button

        'setup mine 
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1
                'locate current button
                MyButton = GetButton(i, j)

                'setup image(mine, number or empty zone)
                If MineMap(i, j) = -1 Then
                    MyButton.Image = ImageMine
                ElseIf MineMap(i, j) = 0 Then
                    MyButton.Image = Nothing
                End If
            Next
        Next
    End Sub

    ' distribute mine marker number
    Private Sub ShowAllMineMarker()
        Dim i, j As Integer
        Dim MyButton As Button

        'setup mine marker number
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1
                'locate current button
                MyButton = GetButton(i, j)

                'setup image(mine, number or empty zone)
                If MineMap(i, j) > 0 Then
                    Set_Button_Mine_Marker(MyButton, MineMap(i, j))
                ElseIf MineMap(i, j) = -9 Then
                    MyButton.Text = ""
                    MyButton.Enabled = False
                End If
            Next
        Next
    End Sub

    ' locate the current button via row and column
    Private Function GetButton(ByVal ColNow, RowNow) As Button
        ' GET BUTTON AND THE DEFAUlT NAME(BUTTON1, BUTTON2, etc.)
        Dim ButtonName As String
        Dim ButtonIndex As Integer
        ButtonIndex = ColNow + 1 + RowNow * 8
        ButtonName = "Button" + CStr(ButtonIndex)
        Return CType(Me.Controls(ButtonName), Button)
    End Function


    Private Sub Button_MouseMove(sender As Object, e As MouseEventArgs) Handles Button96.MouseMove, Button95.MouseMove, Button94.MouseMove, Button93.MouseMove, Button92.MouseMove, Button91.MouseMove, Button90.MouseMove, Button9.MouseMove, Button89.MouseMove, Button88.MouseMove, Button87.MouseMove, Button86.MouseMove, Button85.MouseMove, Button84.MouseMove, Button83.MouseMove, Button82.MouseMove, Button81.MouseMove, Button80.MouseMove, Button8.MouseMove, Button79.MouseMove, Button78.MouseMove, Button77.MouseMove, Button76.MouseMove, Button75.MouseMove, Button74.MouseMove, Button73.MouseMove, Button72.MouseMove, Button71.MouseMove, Button70.MouseMove, Button7.MouseMove, Button69.MouseMove, Button68.MouseMove, Button67.MouseMove, Button66.MouseMove, Button65.MouseMove, Button64.MouseMove, Button63.MouseMove, Button62.MouseMove, Button61.MouseMove, Button60.MouseMove, Button6.MouseMove, Button59.MouseMove, Button58.MouseMove, Button57.MouseMove, Button56.MouseMove, Button55.MouseMove, Button54.MouseMove, Button53.MouseMove, Button52.MouseMove, Button51.MouseMove, Button50.MouseMove, Button5.MouseMove, Button49.MouseMove, Button48.MouseMove, Button47.MouseMove, Button46.MouseMove, Button45.MouseMove, Button44.MouseMove, Button43.MouseMove, Button42.MouseMove, Button41.MouseMove, Button40.MouseMove, Button4.MouseMove, Button39.MouseMove, Button38.MouseMove, Button37.MouseMove, Button36.MouseMove, Button35.MouseMove, Button34.MouseMove, Button33.MouseMove, Button32.MouseMove, Button31.MouseMove, Button30.MouseMove, Button3.MouseMove, Button29.MouseMove, Button28.MouseMove, Button27.MouseMove, Button26.MouseMove, Button25.MouseMove, Button24.MouseMove, Button23.MouseMove, Button22.MouseMove, Button21.MouseMove, Button20.MouseMove, Button2.MouseMove, Button19.MouseMove, Button18.MouseMove, Button17.MouseMove, Button16.MouseMove, Button15.MouseMove, Button14.MouseMove, Button13.MouseMove, Button12.MouseMove, Button11.MouseMove, Button10.MouseMove, Button1.MouseMove
        Dim MyButton As Button = sender
        MyButton.Focus()
    End Sub

    Private Sub Button_MouseUp(sender As Object, e As MouseEventArgs) Handles Button96.MouseUp, Button95.MouseUp, Button94.MouseUp, Button93.MouseUp, Button92.MouseUp, Button91.MouseUp, Button90.MouseUp, Button9.MouseUp, Button89.MouseUp, Button88.MouseUp, Button87.MouseUp, Button86.MouseUp, Button85.MouseUp, Button84.MouseUp, Button83.MouseUp, Button82.MouseUp, Button81.MouseUp, Button80.MouseUp, Button8.MouseUp, Button79.MouseUp, Button78.MouseUp, Button77.MouseUp, Button76.MouseUp, Button75.MouseUp, Button74.MouseUp, Button73.MouseUp, Button72.MouseUp, Button71.MouseUp, Button70.MouseUp, Button7.MouseUp, Button69.MouseUp, Button68.MouseUp, Button67.MouseUp, Button66.MouseUp, Button65.MouseUp, Button64.MouseUp, Button63.MouseUp, Button62.MouseUp, Button61.MouseUp, Button60.MouseUp, Button6.MouseUp, Button59.MouseUp, Button58.MouseUp, Button57.MouseUp, Button56.MouseUp, Button55.MouseUp, Button54.MouseUp, Button53.MouseUp, Button52.MouseUp, Button51.MouseUp, Button50.MouseUp, Button5.MouseUp, Button49.MouseUp, Button48.MouseUp, Button47.MouseUp, Button46.MouseUp, Button45.MouseUp, Button44.MouseUp, Button43.MouseUp, Button42.MouseUp, Button41.MouseUp, Button40.MouseUp, Button4.MouseUp, Button39.MouseUp, Button38.MouseUp, Button37.MouseUp, Button36.MouseUp, Button35.MouseUp, Button34.MouseUp, Button33.MouseUp, Button32.MouseUp, Button31.MouseUp, Button30.MouseUp, Button3.MouseUp, Button29.MouseUp, Button28.MouseUp, Button27.MouseUp, Button26.MouseUp, Button25.MouseUp, Button24.MouseUp, Button23.MouseUp, Button22.MouseUp, Button21.MouseUp, Button20.MouseUp, Button2.MouseUp, Button19.MouseUp, Button18.MouseUp, Button17.MouseUp, Button16.MouseUp, Button15.MouseUp, Button14.MouseUp, Button13.MouseUp, Button12.MouseUp, Button11.MouseUp, Button10.MouseUp, Button1.MouseUp
        ' MOUSE CLICK EVENT(LEFT OF RIGHT)
        Dim PosNow(,), XNow, YNow As Integer
        Dim MyButton As Button

        ' get the current button with row and col number 
        MyButton = sender
        PosNow = GetLocation(MyButton.Name)
        XNow = PosNow(0, 0)
        YNow = PosNow(0, 1)

        If FirstClick = False Then
            FirstClick = True
            Init_Mine_Map(XNow, YNow)
        End If

        'mouse left click
        If e.Button = Windows.Forms.MouseButtons.Left Then
            ' 1. hit a mine, show all mines and markers, game failed
            If MineMap(XNow, YNow) = -1 Then
                ShowAllMine()
                ShowAllMineMarker()

                ' highlight the current mine
                MyButton.Image = ImageHighlight

                ' display failure face
                PictureBox1.Image = ImageFailure

                ' 2. hit a mine marker number, display it
            ElseIf MineMap(XNow, YNow) > 0 Then
                ' show the current mine marker number
                Set_Button_Mine_Marker(MyButton, MineMap(XNow, YNow))

                ' 3. hit empty island spot, display the current island
            ElseIf MineMap(XNow, YNow) = -9 Then
                'show empty space
                ShowSingleIsland(XNow, YNow)
            End If

            ' mouse right click
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            If MyButton.Text = "" Then
                ' 1. right click to setup mine flag
                If MyButton.Image Is Nothing Then
                    'mark the mine with flag
                    MyButton.Image = ImageFlag

                    'count potential remaining mine
                    MineNumLeftDisplay -= 1

                    'count the correct remaining mine
                    If MineMap(XNow, YNow) = -1 Then
                        MineNumLeft -= 1
                    End If
                Else
                    'unmark the mine
                    MyButton.Image = Nothing

                    ' decount potential remaining mine
                    MineNumLeftDisplay += 1

                    'decount the correct remaining mine
                    If MineMap(XNow, YNow) = -1 Then
                        MineNumLeft += 1
                    End If
                End If

                'update current mine recorder
                Label2.Text = CStr(MineNumLeftDisplay) & "/" & CStr(MineNum)
            End If
        End If
        ' display success face
        If MineNumLeft = 0 Then
            PictureBox1.Image = ImageSuccess
        End If
    End Sub

    ' get location of the current button
    Private Function GetLocation(ByRef CurrentName As String) As Integer(,)
        Dim Position(0, 1) As Integer
        Dim ButtonIndex As String = ""
        For Each c As Char In CurrentName
            If IsNumeric(c) Then
                ButtonIndex &= c
            End If
        Next

        Position(0, 0) = ((CInt(ButtonIndex) - 1) Mod MineMapWidth)
        Position(0, 1) = (CInt(ButtonIndex) - 1) \ MineMapWidth
        Return Position
    End Function



    ' get empty islands
    Private Sub GetEmptyIsland()
        Dim i, j As Integer

        'initialize parameters
        TempMineMap = MineMap
        IslandNum = 0
        IslandMap = New Dictionary(Of Integer, Integer(,))

        'search island
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1
                If TempMineMap(i, j) = 0 Then
                    IslandPointNum = 0
                    GetSingleIsland(i, j)
                    IslandNum += 1
                    IslandMap.Add(IslandNum, SingleIslandMap)
                End If
            Next
        Next
    End Sub

    ' find single island
    Private Sub GetSingleIsland(ByVal i As Integer, ByVal j As Integer)
        'check if the position is outside of the margin
        If i >= 0 And i <= MineMapWidth - 1 And j >= 0 And j <= MineMapHeight - 1 Then
            ReDim Preserve SingleIslandMap(1, IslandPointNum)
            SingleIslandMap(0, IslandPointNum) = i
            SingleIslandMap(1, IslandPointNum) = j
            IslandPointNum += 1

            'check unknown island
            If TempMineMap(i, j) = 0 Then
                'elimate and record it
                TempMineMap(i, j) = -9

                'recurse and search the next position
                GetSingleIsland(i - 1, j)
                GetSingleIsland(i + 1, j)
                GetSingleIsland(i, j - 1)
                GetSingleIsland(i, j + 1)
            End If
        End If
    End Sub

    ' show single island
    Private Sub ShowSingleIsland(ByVal i As Integer, ByVal j As Integer)
        Dim IslandIndex, XNow, YNow As Integer
        Dim MyButton As Button
        For Each pair As KeyValuePair(Of Integer, Integer(,)) In IslandMap
            Dim n, m As Integer
            For n = 0 To (pair.Value.Length / 2) - 1
                If pair.Value(0, n) = i And pair.Value(1, n) = j Then
                    IslandIndex = pair.Key
                    For m = 0 To (pair.Value.Length / 2) - 1
                        XNow = IslandMap(IslandIndex)(0, m)
                        YNow = IslandMap(IslandIndex)(1, m)
                        MyButton = GetButton(XNow, YNow)
                        If MineMap(XNow, YNow) = -9 Then
                            MyButton.Text = ""
                            MyButton.Enabled = False
                        ElseIf MineMap(XNow, YNow) > 0 Then
                            Set_Button_Mine_Marker(MyButton, MineMap(XNow, YNow))
                        End If

                    Next
                    Exit Sub
                End If
            Next
        Next
    End Sub

    ' reset all button
    Private Sub InitAllButton()
        Dim i, j As Integer
        Dim MyButton As Button

        'setup mine 
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1
                'locate current button
                MyButton = GetButton(i, j)

                'reset the button
                MyButton.Text = ""
                MyButton.Image = Nothing
                MyButton.Enabled = True
            Next
        Next
    End Sub

    Private Sub Set_Button_Mine_Marker(ByVal MyButton As Button, MarkerNum As Integer)
        MyButton.Font = New Font(MyButton.Font, FontStyle.Bold)
        MyButton.Text = CStr(MarkerNum)

        Select Case MarkerNum
            Case 1
                MyButton.ForeColor = Color.Blue
            Case 2
                MyButton.ForeColor = Color.Green
            Case 3
                MyButton.ForeColor = Color.Red
            Case 4
                MyButton.ForeColor = Color.DarkBlue
            Case Else
                MyButton.ForeColor = Color.Chocolate
        End Select
    End Sub
End Class
