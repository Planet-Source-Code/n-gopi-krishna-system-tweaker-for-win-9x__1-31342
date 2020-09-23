VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "System Tweaking Helper"
   ClientHeight    =   5070
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   8025
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   3840
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4815
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   7575
   End
   Begin VB.OptionButton optPrinters 
      Caption         =   "Users can't add/delete Printers"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   2175
   End
   Begin VB.OptionButton optNoNetHood 
      Caption         =   "No Network  Icon"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.OptionButton optNoDrives 
      Caption         =   "Hide all drives except C:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton optNoFind 
      Caption         =   "No Find Files"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optNoRecentDocs 
      Caption         =   "No Recent Docs menu"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox SysTweak 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "MyCpl.frx":0000
      ToolTipText     =   "Paste this text in a   .REG   file.   Double Click it.   Enjoy!!!"
      Top             =   960
      Width           =   4935
   End
   Begin VB.OptionButton optNoRun 
      Caption         =   "No Run in Start Menu"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton optNoLogOff 
      Caption         =   "No Log Off in Start Menu"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton optNoFav 
      Caption         =   "No Favourites in Start Menu"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.FileListBox cplFileList 
      Height          =   1650
      Left            =   480
      Pattern         =   "*.cpl"
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   697
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "News Gothic MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Control Panel"
      TabPicture(0)   =   "MyCpl.frx":0009
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "System Tweak"
      TabPicture(1)   =   "MyCpl.frx":0025
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "MyCpl.frx":0041
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (lpBuffer As MEMORYSTATUS)


Private Sub cplFileList_DblClick()
Dim command As String
    
    command = "control.exe " + cplFileList.List(cplFileList.ListIndex)
    Shell command, vbMaximizedFocus
End Sub


Private Sub optNoDrives_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoDrives" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoFav_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoFavoritesMenu" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoFind_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoFind" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoLogOff_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoLogOff" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoNetHood_Click()

    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoNetHood" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoRecentDocs_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoRecentDocsMenu" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optNoRun_Click()
    
    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoRun" + Chr(34) + "=dword:00000001"
End Sub

Private Sub optPrinters_Click()

    SysTweak.Text = "REGEDIT4" & vbCrLf & vbCrLf
    SysTweak.Text = SysTweak.Text + "[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" + vbCrLf
    SysTweak.Text = SysTweak.Text + Chr(34) + "NoAddPrinter" + Chr(34) + "=dword:00000001"
    SysTweak.Text = SysTweak.Text + vbCrLf + Chr(34) + "NoDeletePrinter" + Chr(34) + "=dword:00000001"
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim currenttab As Integer
    
    cplFileList.Visible = False
    optNoFav.Visible = False
    optNoRun.Visible = False
    optNoLogOff.Visible = False
    optNoRecentDocs.Visible = False
    optNoFind.Visible = False
    optNoDrives.Visible = False
    optNoNetHood.Visible = False
    optPrinters.Visible = False
    SysTweak.Visible = False
    txtAbout.Visible = False
    
'
'Get the tab currently selected and make visible the items belonging to it.
'
'
    currenttab = SSTab1.Tab
    If (currenttab = 0) Then
        cplFileList.Visible = True
    End If
    If (currenttab = 1) Then
        optNoFav.Visible = True
        optNoRun.Visible = True
        optNoLogOff.Visible = True
        optNoRecentDocs.Visible = True
        optNoFind.Visible = True
        optNoDrives.Visible = True
        optNoNetHood.Visible = True
        optPrinters.Visible = True
        SysTweak.Visible = True
    End If
    If (currenttab = 2) Then
        txtAbout.Visible = True
        txtAbout.Text = "N.Gopi Krishna                             ngopikrishna81@yahoo.com" + vbCrLf
        txtAbout.Text = txtAbout.Text + "4/4 B.Tech C.S & S.E" + vbCrLf
        txtAbout.Text = txtAbout.Text + "A.U.College of Engineering" + vbCrLf
        txtAbout.Text = txtAbout.Text + "Andhra University" + vbCrLf
        txtAbout.Text = txtAbout.Text + vbCrLf + vbCrLf + vbCrLf + vbCrLf
        txtAbout.Text = txtAbout.Text + "In System Tweak tab, you will be shown"
        txtAbout.Text = txtAbout.Text + "a certain text in the text box.Copy it "
        txtAbout.Text = txtAbout.Text + "and create file with extension [.REG]  Paste the text and save the file. "
        txtAbout.Text = txtAbout.Text + "Double click it. Answer YES to the questions asked." + vbCrLf + vbCrLf
        txtAbout.Text = txtAbout.Text + "If you already have done this and want to undo it,"
        txtAbout.Text = txtAbout.Text + "then change the last one of the dword:00000001 to zero"
        txtAbout.Text = txtAbout.Text + "i.e dword:00000000  .Paste the code in a file and repeat as above."
        txtAbout.Text = txtAbout.Text + vbCrLf + vbCrLf + "Some of these options need a restart."
        txtAbout.Text = txtAbout.Text + vbCrLf + vbCrLf + "              All the Best. Do mail your comments."
        
        
                        
                        
                        
                        
    End If
End Sub

Private Sub Form_Load()



    cplFileList.Path = Environ("winbootdir") + "\system"
    cplFileList.Visible = True
    optNoFav.Visible = False
    optNoRun.Visible = False
    optNoLogOff.Visible = False
    optNoFind.Visible = False
    optNoRecentDocs.Visible = False
    optNoDrives.Visible = False
    optNoNetHood.Visible = False
    optPrinters.Visible = False
    SysTweak.Visible = False
    txtAbout.Visible = False
    txtAbout.Height = 4000  'These two lines are for aesthetic reasons.
    txtAbout.Width = 7500   'Just to avoid confusion during coding
End Sub

Private Sub Timer1_Timer()
Dim ms As MEMORYSTATUS
    GlobalMemoryStatus ms
    StatusBar.SimpleText = "Total Physical Memory : "
    StatusBar.SimpleText = StatusBar.SimpleText + Str(ms.dwTotalPhys / 1024) + "KB"
    StatusBar.SimpleText = StatusBar.SimpleText + "               Free Physical Memory : " + Str(ms.dwAvailPhys / 1024) + "KB"
End Sub
