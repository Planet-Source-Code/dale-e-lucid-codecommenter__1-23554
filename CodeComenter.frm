VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FCodeComenter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Commenter"
   ClientHeight    =   5205
   ClientLeft      =   3150
   ClientTop       =   2190
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabMain 
      Height          =   4965
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   8758
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "CodeComenter.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPurpose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAuthor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPurpose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAuthor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "framAS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ckAllStatic"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "framScope"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "framType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CancelButton"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "OKButton"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Custom"
      TabPicture(1)   =   "CodeComenter.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemove"
      Tab(1).Control(1)=   "cmdAdd"
      Tab(1).Control(2)=   "lstCustom"
      Tab(1).Control(3)=   "txtAdd"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "CodeComenter.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPreview"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtAdd 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   -74865
         TabIndex        =   29
         Top             =   3915
         Width           =   4290
      End
      Begin VB.ListBox lstCustom 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   3435
         ItemData        =   "CodeComenter.frx":0054
         Left            =   -74865
         List            =   "CodeComenter.frx":0056
         TabIndex        =   28
         Top             =   405
         Width           =   4275
      End
      Begin VB.TextBox txtPreview 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   4455
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   27
         Top             =   405
         Width           =   4290
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73170
         TabIndex        =   26
         Top             =   4365
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71850
         TabIndex        =   25
         Top             =   4365
         Width           =   1215
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3030
         TabIndex        =   21
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3030
         TabIndex        =   20
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   990
         TabIndex        =   19
         ToolTipText     =   "200"
         Top             =   540
         Width           =   1815
      End
      Begin VB.Frame framType 
         Caption         =   "Type:"
         Height          =   1215
         Left            =   270
         TabIndex        =   14
         ToolTipText     =   "202"
         Top             =   1500
         Width           =   2535
         Begin VB.OptionButton optType 
            Caption         =   "Sub"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Function"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Property"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Event"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame framScope 
         Caption         =   "Scope:"
         Height          =   735
         Left            =   270
         TabIndex        =   11
         ToolTipText     =   "203"
         Top             =   2820
         Width           =   2535
         Begin VB.OptionButton optScope 
            Caption         =   "Public"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optScope 
            Caption         =   "Private"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CheckBox ckAllStatic 
         Caption         =   "&All Local variables as Statics"
         Enabled         =   0   'False
         Height          =   255
         Left            =   270
         TabIndex        =   10
         ToolTipText     =   "204"
         Top             =   3660
         Width           =   2535
      End
      Begin VB.Frame framAS 
         Caption         =   "AS:"
         Height          =   2415
         Left            =   3030
         TabIndex        =   3
         ToolTipText     =   "206"
         Top             =   1500
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optAS 
            Caption         =   "Boolean"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAS 
            Caption         =   "Long"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optAS 
            Caption         =   "String"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton optAS 
            Caption         =   "Double"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton optAS 
            Caption         =   "Variant"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   2040
            Width           =   975
         End
         Begin VB.OptionButton optAS 
            Caption         =   "Intger"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox txtAuthor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   2
         ToolTipText     =   "201"
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtPurpose 
         Enabled         =   0   'False
         Height          =   495
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "CodeComenter.frx":0058
         ToolTipText     =   "205"
         Top             =   4260
         Width           =   3975
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Name:"
         Height          =   255
         Left            =   270
         TabIndex        =   24
         Top             =   570
         Width           =   615
      End
      Begin VB.Label lblAuthor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Author:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   270
         TabIndex        =   23
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblPurpose 
         BackStyle       =   0  'Transparent
         Caption         =   "Purpose:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   270
         TabIndex        =   22
         ToolTipText     =   "206"
         Top             =   4020
         Width           =   615
      End
   End
End
Attribute VB_Name = "FCodeComenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect


Private Function GetInits(sInStr As String) As Boolean
Dim sNewStr         As String
Dim idx             As Integer

    On Error GoTo ErrGetInits

    GetInits = False

    If Len(Trim(sInStr$)) = 0 Then sInStr$ = "   ": Exit Function

    sNewStr$ = Left(sInStr$, 1)
    For idx% = 1 To Len(Trim(sInStr$))
        If Mid$(sInStr$, idx%, 1) = " " Then
            sNewStr$ = sNewStr$ & Mid$(sInStr$, idx% + 1, 1)
        End If
    Next idx%
    
    sInStr$ = sNewStr$

    sInStr$ = sInStr$ & Space(3)
    sInStr$ = Left(sInStr$, 3)
    
    GetInits = True

ExitGetInits:
    Do While Len(sInStr$) > 3
        sInStr$ = Left(sInStr$, Len(sInStr$) - 1)
    Loop
    
    Exit Function

ErrGetInits:
    Resume ExitGetInits:

End Function


Private Function FComm(sInStr As String) As Boolean
    On Error GoTo ErrFComm
    
    Dim lStart          As Long
    Dim idx             As Long


    FComm = False

    If Len(Trim(sInStr$)) = 0 Then Exit Function

    sInStr$ = Replace(sInStr$, Chr(10), "")
    sInStr$ = Replace(sInStr$, Chr(13), "")

    If Len(sInStr$) > 62 Then
        lStart& = 62
        For lStart& = lStart& To 1 Step -1
            If Mid(sInStr$, lStart&, 1) = " " Then
                sInStr$ = Mid(sInStr$, 1, lStart&) & vbCrLf & "'            " & Mid(sInStr$, lStart&)
                lStart& = lStart& + 13
                Exit For
            End If
        Next lStart
      
        Do While Len(Mid(sInStr, lStart)) > 48
            lStart& = lStart + 48
            For lStart& = lStart& To 1 Step -1
                If Mid(sInStr$, lStart&, 1) = " " Then
                    sInStr$ = Mid(sInStr$, 1, lStart&) & vbCrLf & "'            " & Mid(sInStr$, lStart&)
                    lStart& = lStart& + 13
                    Exit For
                End If
            
            Next lStart&
        
        Loop
    End If

    FComm = True

ExitFComm:
    Exit Function

ErrFComm:
    Resume ExitFComm

End Function


Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub cmdAdd_Click()

    If lstCustom.ListCount = 25 Then
        MsgBox "The maximum custom lines allowed is 25.  " & _
               "To add more lines you must first remove some.", _
               vbExclamation + vbOKOnly, "Max Lines Reached!"
    Else
        lstCustom.AddItem Trim(txtAdd.Text)
        txtAdd.Text = ""
    End If
End Sub

Private Sub cmdRemove_Click()
    lstCustom.RemoveItem lstCustom.ListIndex
    cmdRemove.Enabled = False
End Sub


Private Sub Form_Load()
Dim idx         As Integer

    ' Remember last author.
    txtAuthor.Text = GetSetting(App.Title, "Settings", "Author", "")
    
    ' Remember custom header lines.
    For idx% = 0 To 24
        If Trim(GetSetting(App.Title, "Settings", "CL" & Str(idx%), "")) <> "" Then
            lstCustom.AddItem Trim(GetSetting(App.Title, "Settings", "CL" & Str(idx%), ""))
        End If
    Next idx%
    
    ' Set icon from res file.
    Me.Icon = LoadResPicture(9000, 1)
    
    ' Set tool tips from res file
    txtName.ToolTipText = LoadResString(200)
    txtAuthor.ToolTipText = LoadResString(201)
    framType.ToolTipText = LoadResString(202)
    For idx% = 0 To optType.UBound
        optType(idx%).ToolTipText = LoadResString(202)
    Next idx%
    framScope.ToolTipText = LoadResString(203)
    For idx% = 0 To optScope.UBound
        optScope(idx%).ToolTipText = LoadResString(203)
    Next idx%
    ckAllStatic.ToolTipText = LoadResString(204)
    txtPurpose.ToolTipText = LoadResString(205)
End Sub


Private Sub lblAuthor_Click()
    txtAuthor.SetFocus
End Sub


Private Sub lblName_Click()
    txtName.SetFocus
End Sub


Private Sub lblPurpose_Click()
    txtPurpose.SetFocus
End Sub


Private Sub lstCustom_Click()
    cmdRemove.Enabled = True
End Sub

Private Sub OKButton_Click()
Dim sProcStr            As String
Dim sProcBody1          As String
Dim sProcBody2          As String
Dim iRetVal             As Integer
Dim idx                 As Integer
Dim sScopeStr           As String
Dim sAuthInits          As String
Dim sPurpose            As String

    ' Save the author's name for the next time.
    SaveSetting App.Title, "Settings", "Author", txtAuthor.Text

    ' Save the custom header lines.
    For idx% = 0 To 24
        If idx% < lstCustom.ListCount Then
            SaveSetting App.Title, "Settings", "CL" & Str(idx%), Trim(lstCustom.List(idx%))
        Else
            SaveSetting App.Title, "Settings", "CL" & Str(idx%), ""
        End If
    Next idx%
    
    ' If the proc exists then err and exit.
    For idx% = 0 To optType.Count - 1
        If optType(idx%) Then Exit For
    Next idx%

    If VBInstance.ActiveCodePane.CodeModule.Find(optType(idx%).Caption & " " & txtName.Text, 1, 1, -1, -1) Then
        MsgBox "Porcedure " & txtName.Text & " already exists.", vbCritical, "Code Comenter Error"
        Call CancelButton_Click
        Exit Sub
    End If

    
    ' Build the header.
    sProcStr$ = GetHeader

    ' set pre "Exit" string
    sProcBody1$ = sProcBody1$ & Chr(9) & "On Error Goto " & txtName.Text & "_Error" & vbCrLf
    sProcBody1$ = sProcBody1$ & vbCrLf & Chr(9) & "'~~ Local Variables" & vbCrLf & vbCrLf
    sProcBody1$ = sProcBody1$ & vbCrLf & Chr(9) & "'~~ Start of " & txtName.Text & vbCrLf

    ' set post "Exit" string
    sProcBody2$ = sProcBody2$ & vbCrLf & txtName.Text & "_Error:"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'~~ You must have the LCIErrors.dll registered"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'   and you must reference Cahners Business Information"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'   Error Object in your project."
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'~~ Uncoment the next 5 lines to use LCIErrors"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'DIM oErr As New LCIErrors"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'DIM iRetVal As Integer"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'With oErr"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & Chr(9) & "'.Action = lcierr_MsgBox  ' 8 is for a message box."
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & Chr(9) & "'.ErrorNumber = Err.Number"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & Chr(9) & "'.Source = " & Chr(34) & VBInstance.ActiveVBProject.Name & " \ " & VBInstance.SelectedVBComponent.Name & " \ " & txtName.Text & Chr(34)
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & Chr(9) & "'iRetVal% = .TrapErr()"
    sProcBody2$ = sProcBody2$ & vbCrLf & Chr(9) & "'End With" & vbCrLf
    
    ' Scope (Public / Private)
    ' Let's do the Static here too.
    For idx% = 0 To optScope.UBound
        If optScope(idx%) Then
            sScopeStr$ = optScope(idx%).Caption
            sProcStr$ = sProcStr$ & sScopeStr$
            If ckAllStatic Then sProcStr$ = sProcStr$ & " Static "
            Exit For
        End If
    Next idx%

    ' Type (Sub, Function, Property, Event)
    idx% = 0
    For idx% = 0 To optType.UBound
        If optType(idx%) Then Exit For
    Next idx%
    
    Select Case idx%
        ' Sub
        Case 0
            sProcStr$ = sProcStr$ & " Sub " & txtName.Text & "()" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody1$
            sProcStr$ = sProcStr$ & vbCrLf & txtName.Text & "_Exit:"
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Exit Sub" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody2$
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Resume " & txtName.Text & "_Exit"
            sProcStr$ = sProcStr$ & vbCrLf & "End Sub" & vbCrLf
        ' Function
        Case 1
            sProcStr$ = sProcStr$ & " Function " & txtName.Text & "() AS Long" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody1$
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & txtName.Text & " = 0" & vbCrLf
            sProcStr$ = sProcStr$ & vbCrLf & txtName.Text & "_Exit:"
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Exit Function" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody2$
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & txtName.Text & " = Err.Number"
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Resume " & txtName.Text & "_Exit"
            sProcStr$ = sProcStr$ & vbCrLf & "End Function"
        ' Property
        Case 2
            idx% = 0
            For idx% = 0 To optAS.UBound
                If optAS(idx%) Then Exit For
            Next idx%
            ' Property Get
            sProcStr$ = sProcStr$ & " Property Get " & txtName.Text & "() AS " & optAS(idx%).Caption & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody1$
            sProcStr$ = sProcStr$ & vbCrLf & txtName.Text & "_Exit:"
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Exit Property" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody2$
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Resume " & txtName.Text & "_Exit"
            sProcStr$ = sProcStr$ & vbCrLf & "End Property" & vbCrLf & vbCrLf
            ' Property Let
            sProcStr$ = sProcStr$ & sScopeStr$ & " Property Let " & txtName.Text & "(ByVal vNewValue As " & optAS(idx%).Caption & ")" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody1
            sProcStr$ = sProcStr$ & vbCrLf & txtName.Text & "_Exit:"
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Exit Property" & vbCrLf
            sProcStr$ = sProcStr$ & sProcBody2
            sProcStr$ = sProcStr$ & vbCrLf & Chr(9) & "Resume " & txtName.Text & "_Exit"
            sProcStr$ = sProcStr$ & vbCrLf & "End Property" & vbCrLf
        ' Event
        Case 3
            sProcStr$ = sProcStr$ & " Event " & txtName.Text & "()" & vbCrLf
    End Select

    ' Add the procedure to the code window.
    VBInstance.ActiveCodePane.CodeModule.AddFromString sProcStr$
    
    Unload Me
End Sub


Private Sub optType_Click(Index As Integer)
Dim idx         As Integer

    If framAS.Enabled Then
        If optType(Index%).Caption = "Property" Then
            framAS.Visible = True
        Else
            framAS.Visible = False
        End If
        If optType(Index%).Caption = "Event" Then
            ckAllStatic.Enabled = False
            For idx% = 0 To optScope().UBound
                optScope(idx%).Enabled = False
            Next idx%
        Else
            ckAllStatic.Enabled = True
            For idx% = 0 To optScope().UBound
                optScope(idx%).Enabled = True
            Next idx%
        End If
    End If
    
End Sub


Private Sub tabMain_Click(PreviousTab As Integer)
    Select Case tabMain.Tab
        Case 0       '~~ General
        Case 1       '~~ Custom
        Case 2       '~~ Preview
            txtPreview.Text = GetHeader
        Case Else
    End Select

End Sub


Private Sub txtAdd_Change()
    If Trim(txtAdd.Text) = "" Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub

Private Sub txtAuthor_GotFocus()
    With txtAuthor
        .SelStart = 0
        .SelLength = Len(txtAuthor.Text)
    End With
End Sub


Private Sub txtName_Change()
Dim idx         As Integer

    If txtName.Text <> "" Then
        OKButton.Enabled = True
        lblAuthor.Enabled = True
        txtAuthor.Enabled = True
        For idx% = 0 To optType().UBound
            optType(idx%).Enabled = True
        Next idx%
        For idx% = 0 To optScope().UBound
            optScope(idx%).Enabled = True
        Next idx%
        ckAllStatic.Enabled = True
        lblPurpose.Enabled = True
        txtPurpose.Enabled = True
    Else
        OKButton.Enabled = False
        lblAuthor.Enabled = False
        txtAuthor.Enabled = False
        For idx% = 0 To optType().UBound
            optType(idx%).Enabled = False
        Next idx%
        For idx% = 0 To optScope().UBound
            optScope(idx%).Enabled = False
        Next idx%
        ckAllStatic.Enabled = False
        lblPurpose.Enabled = False
        txtPurpose.Enabled = False
    End If

End Sub


Private Sub txtName_GotFocus()
    With txtName
        .SelStart = 0
        .SelLength = Len(txtName.Text)
    End With

End Sub


Private Sub txtPurpose_GotFocus()
    With txtPurpose
        .SelStart = 0
        .SelLength = Len(txtPurpose.Text)
    End With

End Sub


Private Function GetHeader() As String
    Dim sAuthInits         As String
    Dim sPurpose           As String
    Dim sTmpStr            As String

    Dim iRetVal            As Integer
    Dim idx                As Integer

    
    
    ' Guess the Author's Initials.
    sAuthInits$ = txtAuthor.Text
    iRetVal% = GetInits(sAuthInits$)

    ' Format purpose for multi-line.
    If Trim(txtPurpose.Text) = "<Click here and enter a procedure discription>" & vbCrLf Then
        sPurpose$ = ""
    Else
        sPurpose$ = "'    PURPOSE: " & txtPurpose.Text
    End If
    Call FComm(sPurpose$)

    If sPurpose$ <> "" Then sPurpose$ = sPurpose$ & vbCrLf

    ' Comments
    GetHeader = "" & vbCrLf & _
    "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & _
    "'       NAME: " & txtName.Text & vbCrLf & _
    sPurpose$ & _
    "'       DATE: " & Format(Date, "Long Date") & vbCrLf & "'     AUTHOR: " & _
    txtAuthor.Text & vbCrLf & "'  ARGUMENTS: "
    
    ' Custom lines
    If lstCustom.ListCount <> 0 Then
        For idx% = 0 To lstCustom.ListCount - 1
            sTmpStr$ = lstCustom.List(idx%)
            Call FComm(sTmpStr$)
            GetHeader = GetHeader & vbCrLf & sTmpStr$
        Next idx%
    End If
    
    ' Comments cont.
    GetHeader = GetHeader & _
    vbCrLf$ & "'" & _
    vbCrLf & "' VERSION   WHO  DATE       DESCRIPTION" & _
    vbCrLf & "' ========  ===  =========  ==================================" & _
    vbCrLf & "' 00-01-00  " & _
    sAuthInits$ & "  " & Format(Date, "Medium Date") & "  Initial Release" & vbCrLf & "'" & _
    vbCrLf

End Function
