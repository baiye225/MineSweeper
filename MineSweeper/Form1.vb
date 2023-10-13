Public Class Form1
    Private IslandNum, IslandPointNum, MineNumLeft, MineNumLeftDisplay As Integer
    Private PanelWidth, PanelHeight, FormWidth, FormHeight As Integer
    Private CellSize As Integer
    Private FirstClick, isLeftButtonClicked, isRightButtonClicked As Boolean
    Private MyButtons(,) As Button
    Private MineMap(,), TempMineMap(,) As Integer
    Private IslandMap As Dictionary(Of Integer, Dictionary(Of Integer, Integer()))
    Private SingleIslandMap As Dictionary(Of Integer, Integer())
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
        'FORM LOAD

        'initialize all constant parameters
        Set_Constant_Parameters()

        'initialize user interface
        Init_UI()

        'initialize all buttons event
        Init_Button_Move_Event()
        Init_Button_Down_Event()
        Init_Button_Up_Event()

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
        Label2.Text = CStr(MineNum) & "/" & CStr(MineNum)
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

    Private Sub Set_Mine_Map(ByVal X0 As Integer, Y0 As Integer)
        ' DISTRIBUTE MINE MAP

        Dim MineX, MineY As Integer ' position of the current mine
        Dim MineNumNow As Integer = MineNum
        Dim MyNearSpace As Dictionary(Of Integer, Integer())

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
                    'get surrounding mine coordinates
                    MyNearSpace = Get_Near_Space(i, j, "Map", "Mine")

                    'count the surrounding mines
                    MineMap(i, j) = MyNearSpace.Count
                End If
            Next
        Next

        ' initialize island map
        GetEmptyIsland()
    End Sub

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
                AddHandler MyButtons(i, j).MouseMove, AddressOf Button_Mouse_Move
            Next
        Next
    End Sub

    Private Sub Button_Mouse_Move(sender As Object, e As MouseEventArgs)
        ' MOUSE MOVE EVENT
        Dim MyButton As Button = sender
        MyButton.Focus()
    End Sub

    Private Sub Init_Button_Down_Event()
        ' ADD ALL MINE BUTTON UP EVENT
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                AddHandler MyButtons(i, j).MouseDown, AddressOf Button_Mouse_Down
            Next
        Next
    End Sub

    Private Sub Button_Mouse_Down(sender As Object, e As MouseEventArgs)
        ' MOUSE DOWN EVENT
        If e.Button = MouseButtons.Left Then
            isLeftButtonClicked = True
        ElseIf e.Button = MouseButtons.Right Then
            isRightButtonClicked = True
        End If

    End Sub

    Private Sub Init_Button_Up_Event()
        ' ADD ALL MINE BUTTON UP EVENT
        For i As Integer = 0 To MineMapWidth - 1
            For j As Integer = 0 To MineMapHeight - 1
                AddHandler MyButtons(i, j).MouseUp, AddressOf Button_MouseUp
            Next
        Next
    End Sub

    Private Sub Button_MouseUp(sender As Object, e As MouseEventArgs)
        'MOUSE CLICK EVENT(LEFT OF RIGHT)
        Dim x, y As Integer
        Dim MyButton As Button

        'get the current button with row and col number 
        MyButton = sender
        x = MyButton.Name.Split("_")(1)
        y = MyButton.Name.Split("_")(2)

        'set random mine map at the 1st click
        If FirstClick = False Then
            FirstClick = True
            Set_Mine_Map(x, y)
        End If

        'execute different mouse click subroutines
        If isLeftButtonClicked = True And isRightButtonClicked = False Then
            'mouse left click: sweep mine
            Mouse_Left_Click(x, y)

        ElseIf isLeftButtonClicked = False And isRightButtonClicked = True Then
            'mouse right click: add flag
            Mouse_Right_Click(x, y)

        ElseIf isLeftButtonClicked = True And isRightButtonClicked = True Then
            ' mouse double-click: show mine marker around
            Mouse_Left_And_Right_Click(x, y)

        End If

        'reset left and right click marker
        isLeftButtonClicked = False
        isRightButtonClicked = False

        ' display success face
        If MineNumLeft = 0 Then
            PictureBox1.Image = ImageSuccess
        End If
    End Sub


    Private Sub Mouse_Left_Click(ByVal x As Integer, ByVal y As Integer)
        'MOUSE LEFT CLICK: ADD FLAG

        Dim MyButton As Button = MyButtons(x, y)
        If MyButton.Image Is Nothing Then
            If MineMap(x, y) = -1 Then
                ' 1. hit a mine, show all mines, game failed
                ShowAllMine()

                ' highlight the current mine
                MyButton.Image = ImageHighlight

                ' display failure face
                PictureBox1.Image = ImageFailure

            ElseIf MineMap(x, y) > 0 Then
                ' 2. hit a mine marker number, display it
                Set_Button_Mine_Marker(MyButton, MineMap(x, y))

            ElseIf MineMap(x, y) = -9 Then
                ' 3. hit empty island spot, display the current island
                ShowSingleIsland(x, y)
            End If
        End If
    End Sub

    Private Sub Mouse_Right_Click(ByVal x As Integer, ByVal y As Integer)
        'MOUSE RIGHT CLICK: SWEEP MINE

        Dim MyButton As Button = MyButtons(x, y)

        If MyButton.Text = "" Then
            ' 1. right click to setup mine flag
            If MyButton.Image Is Nothing Then
                'mark the mine with flag
                MyButton.Image = ImageFlag

                'count potential remaining mine
                MineNumLeftDisplay -= 1

                'count the correct remaining mine
                If MineMap(x, y) = -1 Then
                    MineNumLeft -= 1
                End If
            Else
                'unmark the mine
                MyButton.Image = Nothing

                ' decount potential remaining mine
                MineNumLeftDisplay += 1

                'decount the correct remaining mine
                If MineMap(x, y) = -1 Then
                    MineNumLeft += 1
                End If
            End If

            'update current mine recorder
            Label2.Text = CStr(MineNumLeftDisplay) & "/" & CStr(MineNum)
        End If
    End Sub

    Private Sub Mouse_Left_And_Right_Click(ByVal x As Integer, ByVal y As Integer)
        'MOUSE LEFT AND RIGHT CLICK SIMULTANEOUSLY: SHOW MINE MARKERS AROUND

        'check if current space is a visible mine number marker
        If MineMap(x, y) > 0 And MyButtons(x, y).Text <> "" Then
            Dim MyNearFlag, MyNearUnclicked As Dictionary(Of Integer, Integer())

            'get surrounding flag and unclicked coordinates
            MyNearFlag = Get_Near_Space(x, y, "User", "Flag")
            MyNearUnclicked = Get_Near_Space(x, y, "User", "Unclicked")

            'check if the number of flag matches the mine marker
            If MyNearFlag.Count = CInt(MyButtons(x, y).Text) Then

                'click all surrouding unclicked space
                For i As Integer = 0 To MyNearUnclicked.Count - 1
                    Dim xi, yi As Integer
                    xi = MyNearUnclicked(i)(0)
                    yi = MyNearUnclicked(i)(1)
                    Mouse_Left_Click(xi, yi)
                Next

            End If

        End If
    End Sub

    Private Function Get_Near_Space(ByVal i As Integer, ByVal j As Integer,
                                    ByVal SpaceType As String,
                                    ByVal MarkerName As String) As Dictionary(Of Integer, Integer())
        'GET THE COORDINATES OF NEAR SPACE WITH GIVING MARKER
        Dim MyNearSpace As New Dictionary(Of Integer, Integer())
        Dim num As Integer
        Dim Vectors(,) As Integer
        Dim iNew, jNew As Integer
        Dim MyStatus As String

        'get 8 surrouding orientations
        Vectors = {{0, -1}, {0, 1}, {-1, 0}, {1, 0},
                  {-1, -1}, {1, -1}, {-1, 1}, {1, 1}}

        'check 8 surrounding orientations
        For n As Integer = 0 To Vectors.Length / 2 - 1
            'get one surrounding space
            iNew = i + Vectors(n, 0)
            jNew = j + Vectors(n, 1)

            'chekc if it exceed the margin
            If iNew >= 0 And iNew <= MineMapWidth - 1 And
               jNew >= 0 And jNew <= MineMapHeight - 1 Then

                'check status
                If SpaceType = "Map" Then
                    'check value in mine map
                    MyStatus = Get_Map_Status(iNew, jNew)

                ElseIf SpaceType = "User" Then
                    'check value in user's map
                    MyStatus = Get_User_Status(iNew, jNew)
                End If

                If MyStatus = MarkerName Then
                    MyNearSpace.Add(num, {iNew, jNew})
                    num += 1
                End If
            End If
        Next

        Return MyNearSpace
    End Function

    Private Function Get_Map_Status(ByVal i As Integer, ByVal j As Integer) As String
        'GET THE STATUS OF THE CURRENT MAP SPACE

        Dim result As String
        Select Case MineMap(i, j)
            Case 0
                result = "Empty"
            Case > 0
                result = "MineMarker"
            Case -1
                result = "Mine"
            Case Else
                result = "Undefined"
        End Select

        Return result
    End Function

    Private Function Get_User_Status(ByVal i As Integer, ByVal j As Integer) As String
        'GET THE STATUS OF THE CURRENT USER SPACE

        Dim result As String

        If MyButtons(i, j).Image Is ImageFlag Then
            result = "Flag"
        ElseIf MyButtons(i, j).Image Is Nothing And
               MyButtons(i, j).Enabled = True And
               MyButtons(i, j).text = "" Then
            result = "Unclicked"
        Else
            result = "Undefined"
        End If

        Return result
    End Function

    Private Sub GetEmptyIsland()
        'GET EMPTY ISLANDS

        Dim i, j As Integer

        'initialize parameters
        TempMineMap = MineMap
        IslandNum = 0
        IslandMap = New Dictionary(Of Integer, Dictionary(Of Integer, Integer()))

        'search island
        For i = 0 To MineMapWidth - 1
            For j = 0 To MineMapHeight - 1
                If TempMineMap(i, j) = 0 Then
                    SingleIslandMap = New Dictionary(Of Integer, Integer())
                    IslandPointNum = 0
                    GetSingleIsland(i, j)
                    IslandMap.Add(IslandNum, SingleIslandMap)
                    IslandNum += 1
                End If
            Next
        Next
    End Sub

    Private Sub GetSingleIsland(ByVal i As Integer, ByVal j As Integer)
        'FIND SINGLE ISLAND

        'check if the position is outside of the margin
        If i >= 0 And i <= MineMapWidth - 1 And j >= 0 And j <= MineMapHeight - 1 Then

            'assign the current location index into the current island(island marker or mine marker)
            SingleIslandMap.Add(IslandPointNum, {i, j})
            IslandPointNum += 1

            'check unknown island
            If TempMineMap(i, j) = 0 Then
                'elimate and record it
                TempMineMap(i, j) = -9

                'recurse and search the next position
                GetSingleIsland(i - 1, j)       'left
                GetSingleIsland(i + 1, j)       'right
                GetSingleIsland(i, j - 1)       'up
                GetSingleIsland(i, j + 1)       'down
                GetSingleIsland(i - 1, j - 1)   'left-up
                GetSingleIsland(i + 1, j - 1)   'right-up
                GetSingleIsland(i + 1, j + 1)   'right-down
                GetSingleIsland(i - 1, j + 1)   'left-down
            End If
        End If
    End Sub

    Private Sub ShowSingleIsland(ByVal i As Integer, ByVal j As Integer)
        'SHOW SINGLE ISLAND
        Dim IslandIndex, XNow, YNow As Integer

        'review each island in the island map
        For Each pair As KeyValuePair(Of Integer, Dictionary(Of Integer, Integer())) In IslandMap

            'review each location in the current island
            For n As Integer = 0 To pair.Value.Count - 1

                'check if current button belong to the current island
                If pair.Value(n)(0) = i And pair.Value(n)(1) = j Then

                    'if matches, confirm the current island index
                    IslandIndex = pair.Key

                    'review all location index of the current island
                    For m As Integer = 0 To pair.Value.Count - 1
                        XNow = IslandMap(IslandIndex)(m)(0)
                        YNow = IslandMap(IslandIndex)(m)(1)

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

    Private Sub Reset_All_Button()
        'RESET ALL BUTTONS
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
