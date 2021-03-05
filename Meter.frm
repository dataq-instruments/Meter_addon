VERSION 5.00
Object = "{38DFB2E2-2748-11D4-ABFF-0040055DE809}#1.0#0"; "WinDaq.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F372358F-A5F3-4F52-804B-38D86867F6B9}#1.0#0"; "xchart.ocx"
Begin VB.Form Form0 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4815
   Icon            =   "Meter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   2  'Point
   ScaleWidth      =   240.75
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XCHARTLib.XChart XChart1 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   360
      Top             =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "009000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   1080
   End
   Begin WINDAQLib.WinDaq WinDaq1 
      Left            =   360
      Top             =   360
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Menu ChannelSetupMenu 
      Caption         =   "&Setup"
      Begin VB.Menu AlwaysOnTopMenu 
         Caption         =   "Always on Top"
      End
      Begin VB.Menu ShowUnitMenu 
         Caption         =   "Show&Unit"
         Checked         =   -1  'True
      End
      Begin VB.Menu CaptionMenu 
         Caption         =   "Ca&ption"
      End
      Begin VB.Menu ModeMenu 
         Caption         =   "Data Window &Mode"
         Begin VB.Menu AverageModeMenu 
            Caption         =   "None"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu AverageModeMenu 
            Caption         =   "Minimum"
            Index           =   1
         End
         Begin VB.Menu AverageModeMenu 
            Caption         =   "Maximum"
            Index           =   2
         End
         Begin VB.Menu AverageModeMenu 
            Caption         =   "Average"
            Index           =   3
         End
         Begin VB.Menu AverageModeMenu 
            Caption         =   "RMS"
            Index           =   4
         End
         Begin VB.Menu AverageModeMenu 
            Caption         =   "AC RMS"
            Index           =   5
         End
      End
      Begin VB.Menu AveragingLengthMenu 
         Caption         =   "Data Window &Length"
      End
      Begin VB.Menu AlarmMenu 
         Caption         =   "&Alarm"
         Begin VB.Menu AlarmEnableMenu 
            Caption         =   "Alarm &Enabled"
         End
         Begin VB.Menu HighLimitMenu 
            Caption         =   "&High Limit (D3)"
         End
         Begin VB.Menu LowLimitMenu 
            Caption         =   "&Low Limit (D2)"
         End
         Begin VB.Menu DigitalAlarm 
            Caption         =   "Digital Flags for Alarm"
         End
      End
      Begin VB.Menu ColorMenu 
         Caption         =   "C&olor"
         Begin VB.Menu ForegroundColorMenu 
            Caption         =   "&Foreground Color"
         End
         Begin VB.Menu BackgroundColorMenu 
            Caption         =   "&Background Color"
         End
         Begin VB.Menu HighLimitColorMenu 
            Caption         =   "&High Limit Color"
         End
         Begin VB.Menu LowLimitColorMenu 
            Caption         =   "Lo&w Limit Color"
         End
      End
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is modified from an example provided by DATAQ

Private Type Size
        cx As Long
        cy As Long
End Type

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim WindaqData As Variant
Dim MaxLimit!
Dim MinLimit!
Dim ChannelNumber%
Dim Channels%
Dim iTemp%
Dim ltemp&
Dim sTemp$
Dim lpSize As Size
Dim HighLimitValue!
Dim LowLimitValue!
Dim temp!
Dim ForegroundColor&
Dim BackgroundColor&
Dim HighLimitColor&
Dim LowLimitColor&
Dim AlarmEnabled As Boolean
Dim ShowUnit As Boolean
Dim AverageMode%
Dim AverageLength%
Dim AlwaysOnTop%
Dim MyCaption As String
Dim LocalCaption As String
Dim b148 As Boolean
Dim bDigital As Boolean


Dim bRunning As Boolean
Dim IniFileName$
Dim IniSection$
Dim KeyName$
Dim KeyValue$
Dim KeySize&
Dim x%
Dim MinValue!
Dim MaxValue!
Dim AvgValue As Double

Const NUMBER_OF_MODES = 5       '1 LESS THAN THE NUMBER OF MODES
Dim FLAG_FILE$
Dim MAX_VAL$

Private Sub AlarmEnableMenu_Click()
    AlarmEnableMenu.Checked = Not AlarmEnableMenu.Checked
    AlarmEnabled = AlarmEnableMenu.Checked
End Sub

Private Sub AlwaysOnTopMenu_Click()
    If AlwaysOnTopMenu.Checked = False Then
        AlwaysOnTopMenu.Checked = True
        AlwaysOnTop = -1
    Else
        AlwaysOnTopMenu.Checked = False
        AlwaysOnTop = -2
    End If
    SetWindowPos Form0.hwnd, AlwaysOnTop, 0, 0, 0, 0, 3
End Sub



Private Sub AverageModeMenu_Click(Index As Integer)
    For x = 0 To NUMBER_OF_MODES
        AverageModeMenu(x).Checked = False
    Next x
        AverageModeMenu(Index).Checked = True
        AverageMode = Index
End Sub

Private Sub AveragingLengthMenu_Click()
    sTemp = InputBox("How many points would you like to have as the data window per channel?", "Data Window Length", AverageLength)
    If Val(sTemp) > 1 Then
        AverageLength = Val(sTemp)
    End If
End Sub

Private Sub BackgroundColorMenu_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.Color = Label1.BackColor
   CommonDialog1.ShowColor
   BackgroundColor = CommonDialog1.Color
   Label1.BackColor = BackgroundColor
   Form0.BackColor = BackgroundColor
ErrHandler:
    Exit Sub
End Sub




Private Sub SaveSetupInfo()
'Caption
    KeyName = "Caption"
    sTemp = LocalCaption
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)

'Foreground Color
    KeyName = "Foreground Color"
    sTemp = ForegroundColor
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Background Color
    KeyName = "Background Color"
    sTemp = BackgroundColor
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'High Limit Color
    KeyName = "High Limit Color"
    sTemp = HighLimitColor
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Low Limit Color
    KeyName = "Low Limit Color"
    sTemp = LowLimitColor
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'High Limit Value
    KeyName = "High Limit Value"
    sTemp = HighLimitValue
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Low Limit Value
    KeyName = "Low Limit Value"
    sTemp = LowLimitValue
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Alarm Enabled
    KeyName = "Alarm Enabled"
    sTemp = AlarmEnabled
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Window Height
    KeyName = "Window Height"
    sTemp = Form0.Height
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Window Width
    KeyName = "Window Width"
    sTemp = Form0.Width
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Window Top
    KeyName = "Window Top"
    sTemp = Str$(Form0.Top)
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Window Left
    KeyName = "Window Left"
    sTemp = Form0.Left
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'Always On Top
    KeyName = "Always On Top"
    sTemp = AlwaysOnTop
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
'ShowUnit
    KeyName = "ShowUnit"
    sTemp = ShowUnit
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)

'AveLength
    KeyName = "AverageLength"
    sTemp = AverageLength
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
    
'AveMode
    KeyName = "AverageMode"
    sTemp = AverageMode
    temp = WritePrivateProfileString(IniSection, KeyName, sTemp, IniFileName)
End Sub




Private Sub DigitalAlarm_Click()
   DigitalAlarm.Checked = Not DigitalAlarm.Checked
End Sub

Private Sub ForegroundColorMenu_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.Color = Label1.ForeColor
   CommonDialog1.ShowColor
   ForegroundColor = CommonDialog1.Color
   Label1.ForeColor = ForegroundColor
   Exit Sub
   
ErrHandler:
    Exit Sub
End Sub


Private Sub Form_Load()
Dim sTempx$
ShowUnit = True
IniFileName = "Meter.ini"
AverageLength = 2

ChannelSetupMenu.Visible = False

bRunning = False

sTempx = Space(20)
ltemp = GetPrivateProfileString("Tools", "Superkey", "0", sTempx, 20, "tpdataq.ini")
XChart1.Key = sTempx


'The format is -DDIxxxNT.DLL -Nxxx -In
If Len(Command$) > 0 Then
    
    sTemp1 = Split(Command$, " ")
    i = UBound(sTemp1, 1)
    If Mid$(sTemp1(0), 1, 1) = "-" And UCase(Mid$(sTemp1(0), 2, 1)) = "D" Then
        WinDaq1.DeviceDriver = Mid$(sTemp1(0), 3, Len(sTemp1(0)) - 2)
    End If
    If i > 0 Then
        If Mid$(sTemp1(1), 1, 1) = "-" And UCase(Mid$(sTemp1(1), 2, 1)) = "N" Then
            Form0.Caption = Replace$(Mid$(sTemp1(1), 3, Len(sTemp1(1)) - 2), "_", " ")
            MyCaption = Replace$(Mid$(sTemp1(1), 3, Len(sTemp1(1)) - 2), "_", " ")
        Else
            MyCaption = Form0.Caption
        End If
        
        If i > 1 Then
            If Mid$(sTemp1(2), 1, 1) = "-" And UCase(Mid$(sTemp1(2), 2, 1)) = "I" Then
                If Val(Mid$(sTemp1(2), 3, Len(sTemp1(2)) - 2)) <> -1 Then
                    ChannelNumber = Val(Mid$(sTemp1(2), 3, Len(sTemp1(2)) - 2))
                Else
                    'No channel selection, so we spawn for all channels!
                    If Len(XChart1.ChartText) > 10 Then
                        MsgBox "Group activation is a bonus for paid registered users only", vbOKOnly, XChart1.ChartText
                    Else
                        ChannelNumber = 0
                        For i = 1 To WinDaq1.ChannelCount - 1
                            a$ = "meter.exe " & "-d" & Format$(WinDaq1.DeviceDriver) & " -n" & Mid$(sTemp1(1), 3, Len(sTemp1(1)) - 2) & " -i" & Format$(i)
                            Shell a$, vbNormalNoFocus
                        Next
                    End If
                End If
            End If
        End If
    End If
End If

IniSection = Format$(MyCaption) & Format$(ChannelNumber + 1)

s$ = WinDaq1.DeviceDriver
If Mid$(s$, 3, 1) = "1" Or Mid$(s$, 3, 1) = "6" Then
    b148 = True
Else
    b148 = False
End If

WinDaq1.Start

If b148 Then
    WinDaq1.DigitalOutput (-32765)
End If

KeyValue = Space(255)
Label1.Move 0, 0, ScaleWidth, ScaleHeight
Label1.FontSize = Label1.Height
GetSetupInfo
MaxLimit = WinDaq1.GetMaxLimit(ChannelNumber)
MinLimit = WinDaq1.GetMinLimit(ChannelNumber)
If WinDaq1.GetPhysicalChannel(ChannelNumber) = -1 Then
    bDigital = True
Else
    bDigital = False
End If

If AlwaysOnTop = -1 Then
    AlwaysOnTopMenu.Checked = True
Else
    AlwaysOnTopMenu.Checked = False
End If
SetWindowPos Form0.hwnd, AlwaysOnTop, 0, 0, 0, 0, 3
End Sub

Private Sub Form_Resize()
Label1.Move 0, 0, ScaleWidth, ScaleHeight
If Label1.Height < Label1.Width / 3 Then
    Label1.FontSize = Label1.Height
Else
    Label1.FontSize = Label1.Width / 3.5
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetupInfo
End Sub

Private Sub HighLimitColorMenu_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.Color = HighLimitColor
   CommonDialog1.ShowColor
   HighLimitColor = CommonDialog1.Color
   Exit Sub
   
ErrHandler:
    Exit Sub
End Sub

Private Sub HighLimitMenu_Click()
    sTemp = InputBox("Input High Limit value.  Must be between " + Str$(MinLimit) + " and " + Str$(MaxLimit), "High Limit Value", HighLimitValue)
    If Val(sTemp) >= MinLimit And Val(sTemp) <= MaxLimit Then
        HighLimitValue = Val(sTemp)
    End If
End Sub


Private Sub Label1_DblClick()
    ChannelSetupMenu.Visible = Not ChannelSetupMenu.Visible
End Sub

Private Sub LowLimitColorMenu_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.Color = LowLimitColor
   CommonDialog1.ShowColor
   LowLimitColor = CommonDialog1.Color
   Exit Sub
   
ErrHandler:
    Exit Sub
End Sub

Private Sub LowLimitMenu_Click()
    sTemp = InputBox("Input Low Limit value. Must be between " + Str$(MinLimit) + " and " + Str$(MaxLimit), "Low Limit Value", LowLimitValue)
    If Val(sTemp) >= MinLimit And Val(sTemp) <= MaxLimit Then
        LowLimitValue = Val(sTemp)
    End If
End Sub



Private Sub ShowUnitMenu_Click()
    ShowUnitMenu.Checked = Not ShowUnitMenu.Checked
    ShowUnit = ShowUnitMenu.Checked
End Sub

Private Sub Timer1_Timer()
Dim MinMean!, MaxMean!, temp&, RMS, dBias As Double
Dim dig As Long

Channels = WinDaq1.ChannelCount
    'AvgValue = WinDaq1.GetScaledData(ChannelNumber)
    If bDigital Then
        WindaqData = WinDaq1.GetDataFrame(1, FormatBinary)
        dig = WindaqData(ChannelNumber, 0) And &HFF00
        'If dig < 0 Then dig = dig + 32768
        dig = dig / 256
        dig = dig And &HFF
        Label1.Caption = Hex$(dig)
        Exit Sub
    Else
        WindaqData = WinDaq1.GetDataFrame(AverageLength, FormatScaled)
    End If
    Select Case AverageMode
     Case 1         'MinMode
       AvgValue = WinDaq1.GetMaxLimit(ChannelNumber)
        For x = 0 To AverageLength - 1
            If AvgValue > WindaqData(ChannelNumber, x) Then AvgValue = WindaqData(ChannelNumber, x)
        Next x
    Case 2          'MaxMode
        AvgValue = WinDaq1.GetMinLimit(ChannelNumber)
        For x = 0 To AverageLength - 1
            If AvgValue < WindaqData(ChannelNumber, x) Then AvgValue = WindaqData(ChannelNumber, x)
        Next x
    Case 3          'AvgMode
        AvgValue = 0
        For x = 0 To AverageLength - 1
            AvgValue = AvgValue + WindaqData(ChannelNumber, x) / AverageLength
        Next x
    Case 4         'RMS Mode
        AvgValue = 0
        RMS = 0
        For x = 0 To AverageLength - 1
            RMS = (RMS - dBias) + (WindaqData(ChannelNumber, x)) ^ 2 / AverageLength
        Next x
        AvgValue = Sqr(RMS)
    Case 5         'AC RMS Mode
        dBias = 0
        AvgValue = 0
        RMS = 0
        'find Bias level (average DC level)
        For x = 0 To AverageLength - 1
            dBias = dBias + WindaqData(ChannelNumber, x)
        Next x
            dBias = dBias / AverageLength
        For x = 0 To AverageLength - 1
            RMS = RMS + ((WindaqData(ChannelNumber, x) - dBias)) ^ 2 / AverageLength
        Next x
        AvgValue = Sqr(RMS)
    Case Else 'None
        AvgValue = WinDaq1.GetScaledData(ChannelNumber)
    End Select
    
    If AlarmEnabled = True Then
        If AvgValue > HighLimitValue Then
            If DigitalAlarm.Checked Then
                WinDaq1.DigitalOutput (8)
            End If
            Label1.ForeColor = HighLimitColor
        ElseIf AvgValue < LowLimitValue Then
            If DigitalAlarm.Checked Then
                WinDaq1.DigitalOutput (4)
            End If
            Label1.ForeColor = LowLimitColor
        Else
            If DigitalAlarm.Checked Then
                If b148 Then
                    WinDaq1.DigitalOutput (-32765)
                End If
                WinDaq1.DigitalOutput (0)
            End If
            Label1.ForeColor = ForegroundColor
        End If
    Else
        Label1.ForeColor = ForegroundColor
    End If
    If ShowUnit Then
        Label1.Caption = Format$(AvgValue, "#0.000") & Format$(WinDaq1.GetUnit(ChannelNumber))
    Else
        Label1.Caption = Format$(AvgValue, "#0.000")
    End If
End Sub
Private Sub CaptionMenu_Click()
    sTemp = InputBox("Enter the Caption for this Meter", , "Channel " + Str$(ChannelNumber))
    If Len(sTemp) > 0 Then
        LocalCaption = sTemp
        Caption = MyCaption + " " + LocalCaption
    End If
End Sub
Public Sub GetSetupInfo()

'Caption
    KeyName = "Caption"
    KeySize = 255
    
    ltemp = GetPrivateProfileString(IniSection, KeyName, "Channel " & Format$(WinDaq1.GetPhysicalChannel(ChannelNumber)), KeyValue, KeySize, IniFileName)
    LocalCaption = KeyValue
    Caption = MyCaption + " " + LocalCaption

'Foreground Color
    KeyName = "Foreground Color"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    ForegroundColor = Val(Left$(KeyValue, ltemp))
'Background Color
    KeyName = "Background Color"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "&H8000000F&", KeyValue, KeySize, IniFileName)
    BackgroundColor = Val(Left$(KeyValue, ltemp))
    Form0.BackColor = BackgroundColor
    Label1.BackColor = BackgroundColor
'High Limit Color
    KeyName = "High Limit Color"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    HighLimitColor = Val(Left$(KeyValue, ltemp))
'Low Limit Color
    KeyName = "Low Limit Color"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    LowLimitColor = Val(Left$(KeyValue, ltemp))
'High Limit Value
    KeyName = "High Limit Value"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    HighLimitValue = Val(Left$(KeyValue, ltemp))
'Low Limit Value
    KeyName = "Low Limit Value"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    LowLimitValue = Val(Left$(KeyValue, ltemp))
'Alarm Enabled
    KeyName = "Alarm Enabled"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "false", KeyValue, KeySize, IniFileName)
    AlarmEnabled = Left$(KeyValue, ltemp)
    AlarmEnableMenu.Checked = AlarmEnabled
    
'AveLength
    KeyName = "AverageLength"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "3", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    AverageLength = temp
    
'AveMode
    KeyName = "AverageMode"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "0", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    
    If temp > NUMBER_OF_MODES Then temp = NUMBER_OF_MODES
    AverageMode = temp
    
    For x = 0 To NUMBER_OF_MODES
        AverageModeMenu(x).Checked = False
    Next x
    AverageModeMenu(AverageMode).Checked = True
    
'ShowUnit
    KeyName = "ShowUnit"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "false", KeyValue, KeySize, IniFileName)
    ShowUnit = Left$(KeyValue, ltemp)
    ShowUnitMenu.Checked = ShowUnit
'Window Height
    KeyName = "Window Height"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "1000", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    Form0.Height = temp
'Window Width
    KeyName = "Window Width"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "3000", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    Form0.Width = temp
'Window Top
    KeyName = "Window Top"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "1000", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    Form0.Top = temp
'Window Left
    KeyName = "Window Left"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "1000", KeyValue, KeySize, IniFileName)
    temp = Val(Left$(KeyValue, ltemp))
    Form0.Left = temp
'Always On Top
    KeyName = "Always On Top"
    KeySize = 255
    ltemp = GetPrivateProfileString(IniSection, KeyName, "-1", KeyValue, KeySize, IniFileName)
    AlwaysOnTop = Val(Left$(KeyValue, ltemp))
End Sub

Private Sub WdqGpsMsg1_WinDaqUpdate()
    Channels = WinDaq1.ChannelCount
End Sub

Private Sub WinDaq1_WinDaqExit()
SaveSetupInfo
End
End Sub
