VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileProcessorForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9930.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   25390
   OleObjectBlob   =   "FileProcessorForm.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "FileProcessorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed
'@Folder FileProcessorForm
Option Explicit

Private totalFiles As Long

'表单初始化
Private Sub UserForm_Initialize()
    ' 初始化窗体
    InitializeForm
    
    MoveWindowOffScreen
    ' 设置随机数生成器
    Randomize
End Sub

' 初始化窗体设计和位置
Private Sub InitializeForm()
    ' 设置窗体基本属性
    Me.Caption = "Excel文件批量处理工具"
    Me.BackColor = RGB(240, 240, 240)
    
    ' 设置窗体大小
    Me.Width = 500
    Me.Height = 400
    
    ' 设置标题标签
    With lblTitle
        .Caption = "Excel文件批量处理工具"
        .Font.Name = "微软雅黑"
        .Font.Bold = True
        .Font.Size = 14
        .ForeColor = RGB(30, 55, 153)
        .BackColor = RGB(240, 240, 240)
        .Top = 15
        .Left = 15
        .Width = 400
        .Height = 30
        .AutoSize = False
    End With
    
    ' 设置文件选择按钮
    With btnSelect
        .Caption = "选择文件"
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Top = 60
        .Left = 15
        .Width = 90
        .Height = 30
        .BackColor = RGB(30, 55, 153)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' 设置文件路径文本框
    With txtFilePath
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .Top = 100
        .Left = 15
        .Width = Me.Width - 40
        .Height = 80
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .BackColor = RGB(250, 250, 250)
        .BorderColor = RGB(200, 200, 200)
    End With
    
    ' 设置开始处理按钮
    With btnStart
        .Caption = "开始处理"
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Top = 190
        .Left = 15
        .Width = 90
        .Height = 30
        .BackColor = RGB(46, 125, 50)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' 设置进度信息标签
    With lblProgressInfo
        .Caption = "处理进度"
        .Font.Name = "微软雅黑"
        .Font.Bold = True
        .Font.Size = 10
        .ForeColor = RGB(50, 50, 50)
        .BackColor = RGB(240, 240, 240)
        .Top = 190
        .Left = 120
        .Width = 200
        .Height = 20
        .AutoSize = False
    End With
    
    ' 设置进度信息文本框
    With txtProgress
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Top = 230
        .Left = 15
        .Width = Me.Width - 40
        .Height = 120
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .BackColor = RGB(250, 250, 250)
        .BorderColor = RGB(200, 200, 200)
    End With
    
    ' 设置状态标签
    With lblStatus
        .Caption = "就绪"
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .ForeColor = RGB(100, 100, 100)
        .BackColor = RGB(240, 240, 240)
        .Top = Me.Height - 40
        .Left = 15
        .Width = 200
        .Height = 20
        .AutoSize = False
    End With
    
    ' 设置关于按钮
    With btnAbout
        .Caption = "关于"
        .Font.Name = "微软雅黑"
        .Font.Size = 9
        .Top = Me.Height - 50
        .Left = Me.Width - 70
        .Width = 50
        .Height = 20
        .BackColor = RGB(150, 150, 150)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' 设置窗体在屏幕中央（多屏幕环境）
    CenterFormOnScreen
End Sub


' 将窗体居中显示在主显示器的工作区域
Private Sub CenterFormOnScreen()
    Me.StartUpPosition = 2 ' CenterScreen
End Sub


' 选择文件按钮点击事件
Private Sub btnSelect_Click()
    Dim fileDialog As fileDialog
    Dim selectedItem As Variant
    Dim filePaths As String
    
    
    ' 创建文件对话框
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .AllowMultiSelect = True
        .Title = "请选择Excel文件"
        .Filters.Add "Excel文件", "*.xlsx; *.xls", 1
        
        ' 如果用户选择了文件并点击了"确定"
        If .Show = -1 Then
            ' 清空当前文件路径
            filePaths = ""
            
            ' 遍历所有选中的文件
            For Each selectedItem In .SelectedItems
                filePaths = filePaths & selectedItem & vbCrLf
            Next selectedItem
            
            ' 更新文件路径文本框（移除最后一个换行符）
            If Len(filePaths) > 0 Then
                filePaths = Left$(filePaths, Len(filePaths) - Len(vbCrLf))
            End If
            txtFilePath.Text = filePaths
            
            ' 更新状态
            totalFiles = UBound(Split(filePaths, vbCrLf)) + 1
            lblStatus.Caption = "已选择 " & totalFiles & " 个文件，等待处理"
        End If
    End With
    
    Set fileDialog = Nothing

End Sub

' 开始处理按钮点击事件
Private Sub btnStart_Click()
    Dim filePaths() As String
    Dim filePath As String
    Dim fileName As String
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim errorMessage As String
    Dim result As String
    Dim originalCalculation As XlCalculation
    Dim timeElapsed As String
    Dim processedFiles As Long
    Dim totalProcessingStartTime As Date
    Dim fileProcessingStartTime As Date
    
    ' 保存当前的计算模式
    originalCalculation = Application.Calculation
    
    ' 清空进度信息
    txtProgress.Text = ""
    
    ' 如果没有选择文件，则退出
    If Trim$(txtFilePath.Text) = "" Then
        MsgBox "请先选择文件！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 分割文件路径
    filePaths = Split(txtFilePath.Text, vbCrLf)
    
    ' 设置处理变量
    totalFiles = UBound(filePaths) + 1

    processedFiles = 0
    totalProcessingStartTime = Now
    
    ' 禁用按钮，防止重复点击
    btnSelect.Enabled = False
    btnStart.Enabled = False
    
    ' 更新状态标签
    lblStatus.Caption = "处理中...(0/" & totalFiles & ")"
    lblStatus.ForeColor = RGB(0, 120, 215)
    Me.Repaint
    
    ' 优化性能设置
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' 确保Excel窗口在屏幕外
    MoveWindowOffScreen
    
    ' 处理每个文件
    For i = 0 To UBound(filePaths)
        filePath = Trim$(filePaths(i))
        processedFiles = i + 1
        fileProcessingStartTime = Now
        
        ' 获取文件名（不含路径）
        fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
        
        ' 更新进度信息和状态
        lblStatus.Caption = "处理中...(" & processedFiles & "/" & totalFiles & ")"
        Me.Repaint
        
        ' 添加文件标题（带编号）
        AppendProgressInfo "【" & processedFiles & "/" & totalFiles & "】处理文件: " & fileName
        
        ' 尝试打开文件
        On Error Resume Next
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        
        ' 检查是否有错误
        If Err.Number <> 0 Then
            ' 记录文件打开错误
            errorMessage = "Error: " & Err.Description
            AppendProgressInfo "  └─ " & errorMessage
            
            Err.Clear
            On Error GoTo 0
        Else
            ' 文件打开成功，处理每个工作表
            On Error GoTo 0
            
            ' 处理工作簿中的每个工作表
            For Each ws In wb.Worksheets
                ' 尝试处理工作表
                On Error Resume Next
                MockFileProc filePath, ws.Name
                
                ' 检查是否有错误
                If Err.Number <> 0 Then
                    ' 记录处理错误
                    result = "Error: " & Err.Description
                    Err.Clear
                Else
                    result = "完成"
                End If
                On Error GoTo 0
                
                ' 直接输出最终状态（不显示ing状态）
                AppendProgressInfo "  ├─ " & ws.Name & "..." & result
            Next ws
            
            ' 显示文件处理时间
            timeElapsed = Format$(Now - fileProcessingStartTime, "hh:mm:ss")
            AppendProgressInfo "  └─ 耗时: " & timeElapsed
            
            ' 关闭工作簿
            wb.Close SaveChanges:=False
        End If
        
        Set wb = Nothing
        
        ' 添加一个空行分隔不同的文件
        If i < UBound(filePaths) Then
            AppendProgressInfo ""
        End If
    Next i
    
    ' 显示总处理时间
    timeElapsed = Format$(Now - totalProcessingStartTime, "hh:mm:ss")
    AppendProgressInfo ""
    AppendProgressInfo "======================="
    AppendProgressInfo "全部处理完成！总耗时: " & timeElapsed
    
    ' 恢复原始设置
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = originalCalculation
    
    ' 更新状态标签
    lblStatus.Caption = "处理完成 - 共 " & totalFiles & " 个文件"
    lblStatus.ForeColor = RGB(46, 125, 50)
    
    ' 启用按钮
    btnSelect.Enabled = True
    btnStart.Enabled = True
    
    ' 提示完成
    MsgBox "所有文件处理完成！" & vbCrLf & "总耗时: " & timeElapsed, vbInformation, "处理完成"
End Sub

' 追加进度信息
Private Sub AppendProgressInfo(info As String)
    ' 追加信息到进度文本框
    txtProgress.Text = txtProgress.Text & info & vbCrLf
    
    ' 滚动到最后一行
    txtProgress.SelStart = Len(txtProgress.Text)
    
    ' 更新界面使用Repaint
    Me.Repaint
End Sub

' 关于按钮点击事件
Private Sub btnAbout_Click()
    MsgBox "Excel文件批量处理工具 v1.0" & vbCrLf & _
           "---------------------------" & vbCrLf & _
           "Copyright ? 2025" & vbCrLf & _
           "一个专业的Excel文件批处理工具，" & vbCrLf & _
           "可以批量打开并处理多个Excel文件。" & vbCrLf & _
           vbCrLf & _
           "使用说明：" & vbCrLf & _
           "1. 点击【选择文件】选择要处理的Excel文件" & vbCrLf & _
           "2. 点击【开始处理】开始批量处理" & vbCrLf & _
           "3. 处理完成后会显示详细结果", _
           vbInformation, "关于"
End Sub

' 表单关闭事件 - 处理表单关闭前的确认
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 如果是用户点击X关闭
    If CloseMode = vbFormControlMenu Then
        ' 显示确认对话框
        Dim response As Long
        response = MsgBox("确定要退出应用程序吗？", vbQuestion + vbYesNoCancel, "确认")
        
        ' 除非明确点击"是"，否则取消关闭
        If response <> vbYes Then
            Cancel = True ' 取消关闭
        End If
    End If
End Sub

' 表单关闭事件
Private Sub UserForm_Terminate()
    ' 关闭工作簿
    ThisWorkbook.Close SaveChanges:=False
    RestoreWindowPosition
End Sub


' 恢复Excel窗口到原始位置
Private Sub RestoreWindowPosition()
    Application.WindowState = xlMaximized
End Sub


' 将Excel窗口移动到屏幕外
Private Sub MoveWindowOffScreen()
    With Application
        .WindowState = xlNormal
        .Left = 10000 ' 移到屏幕左侧外
        .Top = 0
    End With
End Sub


