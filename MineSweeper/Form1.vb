Public Class Form1
    Private IslandNum, IslandPointNum, MineNumLeft, MineNumLeftDisplay As Integer
    Private PanelWidth, PanelHeight, FormWidth, FormHeight As Integer
    Private CellSize As Integer
    Private FirstClick As Boolean = False
    Private MyButtons(,) As Button
    Private MineMap(,), TempMineMap(,), SingleIslandMap(,) As Integer
    Private IslandMap As New Dictionary(Of Integer, Integer(,))
    Private ImageMine, ImageHighlight, ImageFlag,
            ImageNormal, ImageSuccess, ImageFailure As Bitmap

    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        'RESET
        Reset_All_Parameters()
    End Sub

    Private Sub Button_Exit_Click(sender As Object, e As EventArgs) Handles Button_Exit.Click
        'EXIT
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'initialize all constant parameters
        Set_Constant_Parameters()

        'initialize user interface
        Init_UI()

        'initialize all buttons event
        Init_Button_Move_Event()
        Init_Button_Click_Event()

        'generate the whole mine distribution
        'Set_Mine_Map()

        'reset all parameters
        Reset_All_Parameters()
    End Sub

    Private Sub Init_UI()
        'INITIALIZE USER INTERFACE

        'initialize mine map
        Init_Mine_Map()

        'initialize images
        Init_Image()

        'set user interface
        Set_UI_Distribution()
    End Sub

    Private Sub Init_Mine_Map()
        ' INITIALIZE WINDOW AND MINE MAP DISTRIBUTION
        Dim ColNum, RowNum As Integer

        ' set panel size
        ColNum = MineMapWidth
        RowNum = MineMapHeight
        CellSize = 20
        PanelWidth = ColNum * CellSize
        PanelHeight = RowNum * CellSize
        Panel1.Size = New Size(PanelWidth, PanelHeight)

        ' set form size
        FormWidth = PanelWidth + 300
        FormHeight = PanelHeight + 300
        Me.Size = New Size(FormWidth, FormHeight)
        Me.CenterToScreen()

        ' set center alignment of the panel
        Panel1.Location = New Point((FormWidth - PanelWidth) / 2,
                                    (FormHeight - PanelHeight) / 2)

        ' add button on the panel
        ReDim MyButtons(ColNum, RowNum)
        For i As Integer = 0 To ColNum - 1
            For j As Integer = 0 To RowNum - 1
                Dim MyButton As Button = New Button
                MyButton.Size = New Size(CellSize, CellSize)
                MyButton.Location = New Point(i * CellSize, j * CellSize)
                MyButton.Name = "Button" + "_" + CStr(i) + "_" + CStr(j)
                Panel1.Controls.Add(MyButton)
                MyButtons(i, j) = MyButton
            Next
        Next
    End Sub

    Private Sub Init_Image()
        'INITIALIZE ALL IMAGE VARIABLES

        'initialize mine button image
        ImageMine = New Bitmap(New Bitmap(ImageFileMine), CellSize, CellSize)
        ImageFlag = New Bitmap(New Bitmap(ImageFileFlag), CellSize, CellSize)
        ImageHighlight = New Bitmap(New Bitmap(ImageFileHighlight), CellSize, CellSize)

        'initialize face image
        ImageNormal = New Bitmap(ImageFileNormal)
        ImageSuccess = New Bitmap(ImageFileSuccess)
        ImageFailure = New Bitmap(ImageFileFailure)
    End Sub

    Private Sub Set_UI_Distribution()
        'SET UP USER INTERFACE DISTRIBUTION

        'title label
        Label1.Text = "MineSweeper V2.0"
        Label1.Location = New Point((FormWidth - Label1.Width) / 2, 0)

        'notification face
        PictureBox1.Image = ImageNormal
        PictureBox1.SizeMode = 1
        PictureBox1.Location = New Point((FormWidth - PictureBox1.Width) / 2, Label1.Location.Y + 50)

        'mine Status
        Label2.Location = New Point((FormWidth - Label2.Width) / 2, PictureBox1.Location.Y + 50)

        'start button
        Button_Start.Location = New Point(50, FormHeight - Button_Start.Height - 50)

        'exit button
        Button_Exit.Location = New Point(FormWidth - Button_Exit.Width - 50, FormHeight - Button_Exit.Height - 50)
    End Sub

    Private Sub Reset_All_Parameters()
        'RESET ALL PARAMETERS
        'initialize first mine click marker
        FirstClick = False

        'reset mine recorder
        MineNumLeft = MineNum
        MineNumLeftDisplay = MineNum
        Label2.Text = CStr(MineNum) & "/" & CStr(MineNum)

        'reset all buttons
        Reset_All_Button()

        'reset mine map
        ReDim MineMap(MineMapWidth - 1, MineMapHeight - 1)


    End Sub

    'generate the whole distribution
    Private Sub Set_Mine_Map(ByVal X0 As Integer, Y0 As Integer)
        ' DISTRIBUTE MINE MAP
        Dim MineX, MineY As Integer ' position of the current mine
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
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1

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

        ' initialize island map
        GetEmptyIsland()
    End Sub

    ' distribute mine
    Private Sub ShowAllMine()
        ' DISPLAY ALL MINE

        'setup mine 
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                'setup image(mine, number or empty zone)
                If MineMap(i, j) = -1 Then
                    MyButtons(i, j).Image = ImageMine
                ElseIf MineMap(i, j) = 0 Then
                    MyButtons(i, j).Image = Nothing
                End If
            Next
        Next
    End Sub

    ' distribute mine marker number
    Private Sub ShowAllMineMarker()
        ' DISPLAY ALL MINE NUMBER MARKER

        'setup mine marker number
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1

                'setup image(mine, number marker or empty zone)
                If MineMap(i, j) > 0 Then
                    Set_Button_Mine_Marker(MyButtons(i, j), MineMap(i, j))

                ElseIf MineMap(i, j) = -9 Then
                    MyButtons(i, j).Text = ""
                    MyButtons(i, j).Enabled = False
                End If
            Next
        Next
    End Sub

    Private Sub Init_Button_Move_Event()
        ' ADD ALL MINE BUTTON MOVE EVENT
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                AddHandler MyButtons(i, j).MouseMove, AddressOf Button_MouseMove
            Next
        Next
    End Sub

    Private Sub Button_MouseMove(sender As Object, e As MouseEventArgs)
        ' MOUSE MOVE EVENT
        Dim MyButton As Button = sender
        MyButton.Focus()
    End Sub

    Private Sub Init_Button_Click_Event()
        ' ADD ALL MINE BUTTON CLICK EVENT
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                AddHandler MyButtons(i, j).MouseUp, AddressOf Button_MouseUp
            Next
        Next
    End Sub

    Private Sub Button_MouseUp(sender As Object, e As MouseEventArgs)
        ' MOUSE CLICK EVENT(LEFT OF RIGHT)
        Dim XNow, YNow As Integer
        Dim MyButton As Button

        ' get the current button with row and col number 
        MyButton = sender
        XNow = MyButton.Name.Split("_")(1)
        YNow = MyButton.Name.Split("_")(2)

        If FirstClick = False Then
            FirstClick = True
            Set_Mine_Map(XNow, YNow)
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
            'assign the current location index into the current island(island marker or mine marker)
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

        'review each island in the island map
        For Each pair As KeyValuePair(Of Integer, Integer(,)) In IslandMap

            'review each location in the current island
            For n As Integer = 0 To (pair.Value.Length / 2) - 1

                'check if current button belong to the current island
                If pair.Value(0, n) = i And pair.Value(1, n) = j Then

                    'if matches, confirm the current island index
                    IslandIndex = pair.Key

                    'review all location index of the current island
                    For m As Integer = 0 To (pair.Value.Length / 2) - 1
                        XNow = IslandMap(IslandIndex)(0, m)
                        YNow = IslandMap(IslandIndex)(1, m)

                        If MineMap(XNow, YNow) = -9 Then
                            'display island space
                            MyButtons(XNow, YNow).Text = ""
                            MyButtons(XNow, YNow).Enabled = False

                        ElseIf MineMap(XNow, YNow) > 0 Then
                            'display miner marker space around the island
                            Set_Button_Mine_Marker(MyButtons(XNow, YNow),
                                                   MineMap(XNow, YNow))
                        End If

                    Next
                    Exit Sub
                End If
            Next
        Next
    End Sub

    ' reset all button
    Private Sub Reset_All_Button()
        'setup mine 
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                'reset the button
                MyButtons(i, j).Text = ""
                MyButtons(i, j).Image = Nothing
                MyButtons(i, j).Enabled = True
            Next
        Next
    End Sub

    Private Sub Set_Button_Mine_Marker(ByVal MyButton As Button, MarkerNum As Integer)
        'DISPLAY MINE NUMBER MARKER

        'setup bold properties
        MyButton.Font = New Font(MyButton.Font, FontStyle.Bold)

        'setup text color
        Dim MyColor As Color

        'determine color based on the marker number
        Select Case MarkerNum
            Case 1
                MyColor = Color.Blue
            Case 2
                MyColor = Color.Green
            Case 3
                MyColor = Color.Red
            Case 4
                MyColor = Color.DarkBlue
            Case Else
                MyColor = Color.Chocolate
        End Select

        'setup text color
        MyButton.ForeColor = MyColor

        'display text
        MyButton.Text = CStr(MarkerNum)
    End Sub
End Class
