Module Properties
    Public ImageFileMine, ImageFileHighlight, ImageFileFlag,
           ImageFileNormal, ImageFileSuccess, ImageFileFailure As String
    Public MineNum, MineMapWidth, MineMapHeight As Integer

    Public Sub Set_Constant_Parameters()
        'SET UP OVERALL PARAMETERS

        'setup images files' name
        'mine  
        ImageFileMine = "mine.png"                  'normal mine
        ImageFileFlag = "flag.png"                  'mine flag
        ImageFileHighlight = "mine highlighted.png" 'hightlighted mine

        'notified face
        ImageFileNormal = "face normal.jpeg"        'normal
        ImageFileSuccess = "face success.jpeg"      'success
        ImageFileFailure = "face failure.jpeg"      'failure

        'setup mine parameters
        MineNum = 15
        MineMapWidth = 8
        MineMapHeight = 10

    End Sub
End Module
