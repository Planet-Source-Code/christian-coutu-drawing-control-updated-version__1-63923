VERSION 5.00
Object = "*\AObjectDraw.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "*\AColorPalette.vbp"
Begin VB.Form Form1 
   Caption         =   "Object Draw Control Sample"
   ClientHeight    =   9060
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar3 
      Align           =   4  'Align Right
      Height          =   7995
      Left            =   11505
      TabIndex        =   11
      Top             =   750
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   14102
      BandCount       =   2
      Orientation     =   1
      _CBWidth        =   2565
      _CBHeight       =   7995
      _Version        =   "6.7.9782"
      Child1          =   "PicProperty1"
      MinWidth1       =   3195
      MinHeight1      =   2505
      Width1          =   3195
      NewRow1         =   0   'False
      Child2          =   "PicProperty2"
      MinWidth2       =   1530
      MinHeight2      =   2325
      Width2          =   9000
      NewRow2         =   0   'False
      Begin VB.PictureBox PicProperty2 
         BorderStyle     =   0  'None
         Height          =   4380
         Left            =   120
         ScaleHeight     =   4380
         ScaleWidth      =   2325
         TabIndex        =   20
         Top             =   3585
         Width           =   2325
         Begin VB.VScrollBar VScroll3 
            Height          =   285
            Left            =   1920
            Max             =   1
            Min             =   200
            TabIndex        =   43
            Top             =   2100
            Value           =   25
            Width           =   255
         End
         Begin VB.TextBox TxtRound 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            TabIndex        =   42
            Text            =   "25"
            Top             =   2100
            Width           =   405
         End
         Begin VB.TextBox TxtPoint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "3"
            ToolTipText     =   "Border Size"
            Top             =   1500
            Width           =   405
         End
         Begin VB.VScrollBar VScroll2 
            Height          =   285
            Left            =   1920
            Max             =   3
            Min             =   30
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1500
            Value           =   3
            Width           =   255
         End
         Begin VB.ComboBox CboFill 
            Height          =   315
            ItemData        =   "Form1.frx":1601A
            Left            =   900
            List            =   "Form1.frx":16036
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   450
            Width           =   1275
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   285
            Left            =   1890
            Max             =   0
            Min             =   100
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox TxtBorder 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "1"
            ToolTipText     =   "Border Size"
            Top             =   60
            Width           =   405
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   345
            Left            =   60
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   609
            _Version        =   393216
            LargeChange     =   1
            Max             =   360
            SelectRange     =   -1  'True
            TickStyle       =   3
         End
         Begin VB.Label Label5 
            Caption         =   "Round Rectangle Size:"
            Height          =   405
            Left            =   90
            TabIndex        =   41
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Points Qty:"
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   1530
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Fill style:"
            Height          =   255
            Left            =   90
            TabIndex        =   27
            Top             =   510
            Width           =   645
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Border size:"
            Height          =   255
            Left            =   90
            TabIndex        =   26
            Top             =   90
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rotation: 0°"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   25
            Top             =   870
            Width           =   1995
         End
      End
      Begin VB.PictureBox PicProperty1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3195
         Left            =   30
         ScaleHeight     =   3195
         ScaleWidth      =   2505
         TabIndex        =   12
         Top             =   165
         Width           =   2505
         Begin VB.VScrollBar ScrCol 
            Height          =   285
            Index           =   2
            Left            =   2100
            Max             =   0
            Min             =   255
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2355
            Width           =   225
         End
         Begin VB.VScrollBar ScrCol 
            Height          =   285
            Index           =   1
            Left            =   1350
            Max             =   0
            Min             =   255
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   2355
            Width           =   225
         End
         Begin VB.VScrollBar ScrCol 
            Height          =   285
            Index           =   0
            Left            =   570
            Max             =   0
            Min             =   255
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2355
            Width           =   225
         End
         Begin VB.TextBox TxtColor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   2340
            Width           =   465
         End
         Begin VB.TextBox TxtColor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "0"
            Top             =   2340
            Width           =   465
         End
         Begin VB.TextBox TxtColor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   2340
            Width           =   465
         End
         Begin VB.OptionButton OpColor 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   1350
            MouseIcon       =   "Form1.frx":1608D
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   1530
            Width           =   945
         End
         Begin VB.OptionButton OpColor 
            BackColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   1350
            MouseIcon       =   "Form1.frx":161DF
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   900
            Width           =   945
         End
         Begin VB.OptionButton OpColor 
            BackColor       =   &H00FF0000&
            Height          =   375
            Index           =   0
            Left            =   1365
            MouseIcon       =   "Form1.frx":16331
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   270
            Value           =   -1  'True
            Width           =   945
         End
         Begin ColorPalette.ColorPal ColorPal1 
            Height          =   2040
            Left            =   120
            TabIndex        =   16
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   3598
            Thumbsize       =   6
         End
         Begin VB.Label LblCol 
            BackStyle       =   0  'Transparent
            Caption         =   "Blue:"
            Height          =   225
            Index           =   2
            Left            =   1650
            TabIndex        =   40
            Top             =   2130
            Width           =   675
         End
         Begin VB.Label LblCol 
            BackStyle       =   0  'Transparent
            Caption         =   "Green:"
            Height          =   225
            Index           =   1
            Left            =   870
            TabIndex        =   39
            Top             =   2130
            Width           =   675
         End
         Begin VB.Label LblCol 
            BackStyle       =   0  'Transparent
            Caption         =   "Red:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   2130
            Width           =   675
         End
         Begin VB.Label LblColor 
            Alignment       =   2  'Center
            Height          =   465
            Left            =   60
            TabIndex        =   28
            Top             =   2730
            Width           =   2385
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color"
            Height          =   255
            Index           =   2
            Left            =   1260
            TabIndex        =   19
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Border Color"
            Height          =   255
            Index           =   1
            Left            =   1290
            TabIndex        =   18
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Color"
            Height          =   255
            Index           =   3
            Left            =   1290
            TabIndex        =   17
            Top             =   90
            Width           =   975
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   8745
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Mouse Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6465
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "22:32"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5220
      Top             =   7860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16483
            Key             =   "Select"
            Object.Tag             =   "Select"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16595
            Key             =   "Line"
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":166A7
            Key             =   "Arc"
            Object.Tag             =   "Arc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":167B9
            Key             =   "Rectangle"
            Object.Tag             =   "Rectangle"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":168CB
            Key             =   "RoundRectangle"
            Object.Tag             =   "RoundRectangle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":169DD
            Key             =   "Ellipse"
            Object.Tag             =   "Ellipse"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16AEF
            Key             =   "Polygon"
            Object.Tag             =   "Polygon"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16E41
            Key             =   "Star"
            Object.Tag             =   "Star"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17193
            Key             =   "Text"
            Object.Tag             =   "Text"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":172A5
            Key             =   "Picture"
            Object.Tag             =   "Picture"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   3  'Align Left
      Height          =   7995
      Left            =   0
      TabIndex        =   8
      Top             =   750
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   14102
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   375
      _CBHeight       =   7995
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar4"
      MinHeight1      =   315
      Width1          =   2640
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   3300
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   5821
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Select"
               Object.ToolTipText     =   "Select Object"
               Object.Tag             =   "Select"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line"
               Object.ToolTipText     =   "Draw Line"
               Object.Tag             =   "Line"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Arc"
               Object.ToolTipText     =   "Draw Arc"
               Object.Tag             =   "Arc"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rectangle"
               Object.ToolTipText     =   "Draw Rectangle"
               Object.Tag             =   "Rectangle"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RoundRectangle"
               Object.ToolTipText     =   "Draw Round Rectangle"
               Object.Tag             =   "RoundRectangle"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ellipse"
               Object.ToolTipText     =   "Draw Ellipse"
               Object.Tag             =   "Ellipse"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Polygon"
               Object.ToolTipText     =   "Draw Polygon"
               Object.Tag             =   "Polygon"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Star"
               Object.ToolTipText     =   "Draw Star"
               Object.Tag             =   "Star"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Text"
               Object.ToolTipText     =   "Draw Text"
               Object.Tag             =   "Text"
               ImageIndex      =   9
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Picture"
               Object.ToolTipText     =   "Insert Picture"
               Object.Tag             =   "Picture"
               ImageIndex      =   10
               Style           =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   1323
      _CBWidth        =   14070
      _CBHeight       =   750
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinWidth1       =   9405
      MinHeight1      =   330
      Width1          =   9405
      NewRow1         =   0   'False
      Child2          =   "Toolbar3"
      MinWidth2       =   1200
      MinHeight2      =   330
      Width2          =   7995
      NewRow2         =   0   'False
      Child3          =   "Toolbar2"
      MinWidth3       =   5595
      MinHeight3      =   330
      Width3          =   4005
      NewRow3         =   -1  'True
      AllowVertical3  =   0   'False
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   9795
         TabIndex        =   7
         Top             =   30
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnZoom"
               Object.ToolTipText     =   "UnZoom"
               Object.Tag             =   "UnZoom"
               ImageIndex      =   33
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom-"
               Object.ToolTipText     =   "Zoom -"
               Object.Tag             =   "Zoom-"
               ImageIndex      =   34
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom+"
               Object.ToolTipText     =   "Zoom+"
               Object.Tag             =   "Zoom+"
               ImageIndex      =   35
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   390
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SelectAll"
               Object.ToolTipText     =   "Select All"
               Object.Tag             =   "SelectAll"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnselectAll"
               Object.ToolTipText     =   "Unselect All"
               Object.Tag             =   "UnselectAll"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignLeft"
               Object.ToolTipText     =   "Align Left"
               Object.Tag             =   "AlignLeft"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignCenterVertical"
               Object.ToolTipText     =   "Align Center Vertical"
               Object.Tag             =   "AlignCenterVertical"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignRight"
               Object.ToolTipText     =   "Align Right"
               Object.Tag             =   "AlignRight"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignTop"
               Object.ToolTipText     =   "Align Top"
               Object.Tag             =   "AlignTop"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignCenterHorizontal"
               Object.ToolTipText     =   "Align Center Horizontal"
               Object.Tag             =   "AlignCenterHorizontal"
               ImageIndex      =   24
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignBottom"
               Object.ToolTipText     =   "Align Bottom"
               Object.Tag             =   "AlignBottom"
               ImageIndex      =   25
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AlignCenterHorVert"
               Object.ToolTipText     =   "Align Center Horizontal+Vertical"
               Object.Tag             =   "AlignCenterHorVert"
               ImageIndex      =   26
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BringToFront"
               Object.ToolTipText     =   "Bring to Front"
               Object.Tag             =   "BringToFront"
               ImageIndex      =   27
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SendToBack"
               Object.ToolTipText     =   "Send to Back"
               Object.Tag             =   "SendToBack"
               ImageIndex      =   28
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BringForward"
               Object.ToolTipText     =   "Bring Forward"
               Object.Tag             =   "BringForward"
               ImageIndex      =   29
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SendBackward"
               Object.ToolTipText     =   "Send Backward"
               Object.Tag             =   "SendBackward"
               ImageIndex      =   30
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Group"
               Object.ToolTipText     =   "Group"
               Object.Tag             =   "Group"
               ImageIndex      =   31
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ungroup"
               Object.ToolTipText     =   "Ungroup"
               Object.Tag             =   "Ungroup"
               ImageIndex      =   32
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   23
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               Object.Tag             =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open"
               Object.Tag             =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save"
               Object.Tag             =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Export"
               Object.ToolTipText     =   "Export"
               Object.Tag             =   "Export"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               Object.Tag             =   "Cut"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               Object.Tag             =   "Copy"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               Object.Tag             =   "Paste"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   "Undo"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   "Redo"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
               Object.Tag             =   "Delete"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextLeft"
               Object.ToolTipText     =   "Align Text Left"
               Object.Tag             =   "AlignText"
               ImageIndex      =   11
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextCenter"
               Object.ToolTipText     =   "Align Text Center"
               Object.Tag             =   "AlignText"
               ImageIndex      =   12
               Style           =   2
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextRight"
               Object.ToolTipText     =   "Align Text Right"
               Object.Tag             =   "AlignText"
               ImageIndex      =   13
               Style           =   2
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               Object.Tag             =   "Bold"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               Object.Tag             =   "Italic"
               ImageIndex      =   15
               Style           =   1
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               Object.Tag             =   "Underline"
               ImageIndex      =   16
               Style           =   1
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Strikethru"
               Object.ToolTipText     =   "Strikethru"
               Object.Tag             =   "Strikethru"
               ImageIndex      =   17
               Style           =   1
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox CboFontName 
            Height          =   315
            Left            =   6540
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Font Name"
            Top             =   0
            Width           =   1905
         End
         Begin VB.ComboBox CboFontSize 
            Height          =   315
            IntegralHeight  =   0   'False
            Left            =   8460
            TabIndex        =   4
            Text            =   "15"
            ToolTipText     =   "Font Size"
            Top             =   0
            Width           =   705
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   7770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":175F7
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17709
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1781B
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1792D
            Key             =   "Export"
            Object.Tag             =   "Export"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17C7F
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17D91
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17EA3
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17FB5
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":180C7
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":181D9
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":182EB
            Key             =   "TextLeft"
            Object.Tag             =   "TextLeft"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":183FD
            Key             =   "TextCenter"
            Object.Tag             =   "TextCenter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1850F
            Key             =   "TextRight"
            Object.Tag             =   "TextRight"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18621
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18733
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18845
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18957
            Key             =   "Strikethru"
            Object.Tag             =   "Strikethru"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18A69
            Key             =   "SelectAll"
            Object.Tag             =   "SelectAll"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18DBB
            Key             =   "UnselectAll"
            Object.Tag             =   "UnselectAll"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1910D
            Key             =   "AlignLeft"
            Object.Tag             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1945F
            Key             =   "AlignCenterVertical"
            Object.Tag             =   "AlignCenterVertical"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":197B1
            Key             =   "AlignRight"
            Object.Tag             =   "AlignRight"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19B03
            Key             =   "AlignTop"
            Object.Tag             =   "AlignTop"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19E55
            Key             =   "AlignCenterHorizontal"
            Object.Tag             =   "AlignCenterHorizontal"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A1A7
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A4F9
            Key             =   "AlignCenterVerticalHorizontal"
            Object.Tag             =   "AlignCenterVerticalHorizontal"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A84B
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A95D
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AA6F
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AB81
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AC93
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ADA5
            Key             =   "Ungroup"
            Object.Tag             =   "Ungroup"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AEB7
            Key             =   "Zoom100"
            Object.Tag             =   "Zoom100"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B209
            Key             =   "Zoom-"
            Object.Tag             =   "Zoom-"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B55B
            Key             =   "Zoom+"
            Object.Tag             =   "Zoom+"
         EndProperty
      EndProperty
   End
   Begin ObjectDraw.ObjDraw ObjDraw1 
      Height          =   6900
      Left            =   600
      TabIndex        =   0
      Top             =   930
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   12171
      CanvasWidth     =   800
      CanvasHeight    =   600
      UndoBufferSize  =   20
   End
   Begin VB.PictureBox PicLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   1350
      ScaleHeight     =   525
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   5820
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   0
      Top             =   7980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu SmnuFile 
         Caption         =   "New"
         Index           =   0
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Save"
         Index           =   2
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Export to Bitmap"
         Index           =   3
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Print"
         Index           =   5
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Exit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu SmnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Delete"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Select All"
         Index           =   9
         Shortcut        =   ^A
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Group"
         Index           =   11
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Ungroup"
         Index           =   12
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "Zoom (100%)"
      Begin VB.Menu SmnuZoom 
         Caption         =   "10%"
         Index           =   0
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "25%"
         Index           =   1
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "50%"
         Index           =   2
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "100%"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "150%"
         Index           =   4
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "200%"
         Index           =   5
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "400%"
         Index           =   6
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu SmnuOptions 
         Caption         =   "Canvas Size"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim doNothing As Boolean
Dim Answer As VbMsgBoxResult
Dim Modified As Boolean

Dim mTxtAlign As AlignmentConstants
Dim mBold As Boolean
Dim mItalic As Boolean
Dim mUnderline As Boolean
Dim mStrikethru As Boolean
Dim ColorIndex As Integer

Dim bFillColor As Long
Dim bBorderColor As Long
Dim bBackColor As Long
Dim bPtsQty As Integer
Private Function FileExist(ByVal MyFile As String) As Boolean
    FileExist = (Dir(MyFile) <> "")
End Function
Private Sub CboFill_Click()
If doNothing = True Then Exit Sub
If ObjDraw1.CurrentObject > -1 Then
ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
End If
End Sub


Private Sub CboFontName_Click()
On Error Resume Next
If doNothing = True Then Exit Sub
If ObjDraw1.ObjectType = mText And doNothing = False Then
ObjDraw1.ModifyObject , , , , , , , , , , CboFontName.Text
End If
ObjDraw1.SetFocus
End Sub


Private Sub CboFontSize_Change()
On Error Resume Next

If doNothing = True Then Exit Sub
If ObjDraw1.ObjectType = mText And doNothing = False And Len(Trim(CboFontSize.Text)) > 0 Then
ObjDraw1.ModifyObject , , , , , , , , , , , CboFontSize.Text
End If
ObjDraw1.SetFocus
End Sub

Private Sub CboFontSize_Click()
If ObjDraw1.ObjectType = mText And doNothing = False Then
ObjDraw1.ModifyObject , , , , , , , , , , , CboFontSize.Text
End If
End Sub





Private Sub ColorPal1_ColorOver(cColor As Long)
Dim sTmp As String
sTmp = Right("000000" & Hex(cColor), 6)
LblColor.Caption = "Hex:" & sTmp & vbCrLf & " Red:" & Int("&H" & Right$(sTmp, 2)) & _
" - Green:" & Int("&H" & Mid$(sTmp, 3, 2)) & " - Blue:" & Int("&H" & Left$(sTmp, 2))

End Sub

Private Sub ColorPal1_ColorSelected(cColor As Long)
Dim sTmp As String
sTmp = Right("000000" & Hex(cColor), 6)
ScrCol(0).Value = Int("&H" & Right$(sTmp, 2))
ScrCol(1).Value = Int("&H" & Mid$(sTmp, 3, 2))
ScrCol(2).Value = Int("&H" & Left$(sTmp, 2))

OpColor(ColorIndex).BackColor = cColor
bFillColor = OpColor(0).BackColor
bBorderColor = OpColor(1).BackColor
bBackColor = OpColor(2).BackColor

If doNothing = True Then Exit Sub

Select Case ColorIndex
    Case 0
    If ObjDraw1.CurrentObject > -1 Then
    ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
    End If
    Case 1
    If ObjDraw1.CurrentObject > -1 Then
    ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
    End If
    Case 2
    ObjDraw1.BackColor = bBackColor
End Select
ObjDraw1.SetFocus
End Sub

Private Sub ColorPal1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    With cDialog
    .DialogTitle = "Open Palette"
    .Filter = "Palette (*.pal)|*.pal"
    .FileName = ""
    .ShowOpen
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 And FileExist(.FileName) = True Then
        ColorPal1.LoadPalette .FileName
        End If
    End With
End If
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
Form_Resize
End Sub

Private Sub CoolBar3_HeightChanged(ByVal NewHeight As Single)
Form_Resize
End Sub

Private Sub Form_Load()
Dim n As Integer
CboFontName.Clear

For n = 1 To Screen.FontCount - 1
CboFontName.AddItem Screen.Fonts(n)
Next n

CboFontName.Text = "Arial"

For n = 5 To 100
CboFontSize.AddItem n
Next n

CboFontSize.Text = 15
CboFill.ListIndex = 0
ColorIndex = 0

bFillColor = OpColor(0).BackColor
bBorderColor = OpColor(1).BackColor
bBackColor = OpColor(2).BackColor

End Sub


Private Sub Form_Paint()
UpdToolBar
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 1 Then
ObjDraw1.Width = Me.Width - CoolBar2.Width - CoolBar3.Width - 175
ObjDraw1.Height = Me.Height - CoolBar1.Height - StatusBar1.Height - 890
ObjDraw1.Top = CoolBar1.Height
ObjDraw1.Left = CoolBar2.Width

End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Modified = True Then
Answer = MsgBox("The drawing as been modified, do you want to save-it before quit?", vbDefaultButton1 + vbYesNoCancel)
    If Answer = vbYes Then
    Cancel = True
        With cDialog
        .DialogTitle = "Save Project"
        .Filter = "Object Draw Project File (*.ojp)|*.ojp"
        .FileName = ""
        .ShowSave
        .FileName = Trim(.FileName)
            If Len(.FileName) > 0 Then
            ObjDraw1.SaveProjects .FileName
            End If
        End With
    End
    ElseIf Answer = vbCancel Then
    Cancel = True
    Exit Sub
    End If
End If
End
End Sub


Private Sub LblColor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
LblColor.Caption = ""
End Sub


Private Sub mnuEdit_Click()
If ObjDraw1.CurrentObject > -1 Then
SmnuEdit(3).Enabled = True
SmnuEdit(4).Enabled = True
SmnuEdit(7).Enabled = True
Else
SmnuEdit(3).Enabled = False
SmnuEdit(4).Enabled = False
SmnuEdit(7).Enabled = False
End If

SmnuEdit(5).Enabled = ObjDraw1.ObjectInClipboard
End Sub

Private Sub mnuZoom_Click()
StatusBar1.Panels(1).Text = "You can also change the Zoom Factor with ""+"" & ""-"" on KeyPad"
End Sub

Private Sub ObjDraw1_KeyDown(KeyAscii As Integer, Shift As Integer)
If KeyAscii >= 37 And KeyAscii <= 40 Then
StatusBar1.Panels(1).Text = "Press ""Ctrl"" key with arrows keys to switch selection"
ElseIf KeyAscii = vbKeyAdd Or KeyAscii = vbKeySubtract Then
mnuZoom.Caption = "Zoom (" & Round(ObjDraw1.ZoomFactor * 100) & "%)"
End If
End Sub

Private Sub ObjDraw1_KeyUp(KeyAscii As Integer, Shift As Integer)
StatusBar1.Panels(1).Text = ""
End Sub


Private Sub ObjDraw1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
StatusBar1.Panels(2).Text = "X: " & x & " - Y: " & Y
End Sub

Private Sub ObjDraw1_NewDrawingEnd()
Toolbar4.Buttons(1).Value = tbrPressed
End Sub

Private Sub ObjDraw1_ObjectResize(ObjType As ObjectDraw.myObType, Index As Long, ObjLeft As Single, ObjTop As Single, ObjWidth As Single, ObjHeight As Single, ObjAspect As Single)
Dim tmp As String

Select Case ObjType
    Case mline
    tmp = "Line"
    Case mArc
    tmp = "Arc"
    Case mRectangle
        If ObjAspect = 0 Then
        tmp = "Rectangle"
        Else
        tmp = "Square"
        End If
    Case mEllipse
        If ObjAspect = 0 Then
        tmp = "Ellipse"
        Else
        tmp = "Circle"
        End If
    Case mText
    tmp = "Text"
    Case mImage
    tmp = "Image"
End Select

StatusBar1.Panels(3).Text = tmp & "   Pos. X:" & ObjLeft & "  Y:" & ObjTop & _
"   Size W:" & ObjWidth & "  H:" & ObjHeight & " "

End Sub

Private Sub ObjDraw1_ObjSelected(ObjType As ObjectDraw.myObType, Index As Long, ObjLeft As Single, ObjTop As Single, ObjWidth As Single, ObjHeight As Single, ObjAngle As Single, ObjFillColor As Long, ObjFillStyle As myFill, ObjBorderColor As Long, ObjBorderWidth As Integer, ObjAspect As Single, ObjFontName As String, ObjFontSize As Single, ObjFontBold As Boolean, ObjFontItalic As Boolean, ObjFontUnderline As Boolean, ObjFontStrikethru As Boolean, ObjText As String, ObjTextAlign As AlignmentConstants, ObjPointQty As Integer)
Dim tmp As String

If ObjType <> -1 Then
doNothing = True
    If ObjFillColor > -1 Then bFillColor = ObjFillColor
    OpColor(0).BackColor = bFillColor
    If ObjFillStyle > -1 Then CboFill.ListIndex = ObjFillStyle
    If ObjAngle > -1 Then Slider1.Value = ObjAngle
    Label1(0).Caption = "Rotation: " & Slider1.Value & "°"
    If ObjBorderColor > -1 Then bBorderColor = ObjBorderColor
    OpColor(1).BackColor = bBorderColor
    If ObjType <> mText Then VScroll1.Value = ObjBorderWidth
    TxtBorder.Text = VScroll1.Value
    If ObjFontName <> "" Then CboFontName.Text = ObjFontName
    If ObjFontSize > 0 Then CboFontSize.Text = ObjFontSize
    If ObjType = mPolygon Or ObjType = mStar Then
        If ObjPointQty > 0 And ObjPointQty <= 30 Then
        VScroll2.Value = ObjPointQty
        TxtPoint.Text = ObjPointQty
        End If
    ElseIf ObjType = mRoundRectangle Then
    VScroll3.Value = ObjPointQty
    TxtRound.Text = ObjPointQty
    End If
    TxtPoint.Text = VScroll2.Value
    mBold = CBool(Int(ObjFontBold))
    Toolbar1.Buttons(19).Value = Abs(Int(ObjFontBold))
    mItalic = CBool(Int(ObjFontItalic))
    Toolbar1.Buttons(20).Value = Abs(Int(ObjFontItalic))
    mUnderline = CBool(Int(ObjFontUnderline))
    Toolbar1.Buttons(21).Value = Abs(Int(ObjFontUnderline))
    mStrikethru = CBool(Int(ObjFontStrikethru))
    Toolbar1.Buttons(22).Value = Abs(Int(ObjFontStrikethru))

    If ObjTextAlign > -1 Then
    mTxtAlign = ObjTextAlign
    Select Case mTxtAlign
        Case vbLeftJustify
        Toolbar1.Buttons(15).Value = tbrPressed
        Case vbRightJustify
        Toolbar1.Buttons(17).Value = tbrPressed
        Case vbCenter
        Toolbar1.Buttons(16).Value = tbrPressed
    End Select
    
    End If
doNothing = False

Select Case ObjType
    Case mline
    tmp = "Line"
    Case mArc
    tmp = "Arc"
    Case mRectangle
        If ObjAspect = 0 Then
        tmp = "Rectangle"
        Else
        tmp = "Square"
        End If
    Case mEllipse
        If ObjAspect = 0 Then
        tmp = "Ellipse"
        Else
        tmp = "Circle"
        End If
    Case mText
    tmp = "Text"
    Case mImage
    tmp = "Image"
    Case mPolygon
    tmp = "Polygon"
    Case mStar
    tmp = "Star"
End Select
StatusBar1.Panels(3).Text = tmp & "   Pos. X:" & ObjLeft & "  Y:" & ObjTop & _
"   Size W:" & ObjWidth & "  H:" & ObjHeight & " "
Else
StatusBar1.Panels(3).Text = ""
End If
UpdToolBar

End Sub


Private Sub ObjDraw1_Prompt2Save()
Modified = True
End Sub

Private Sub ObjDraw1_UndoRedo(LastUndo As Boolean, LastRedo As Boolean)
SmnuEdit(0).Enabled = Not LastUndo
SmnuEdit(1).Enabled = Not LastRedo
Toolbar1.Buttons(10).Enabled = Not LastUndo
Toolbar1.Buttons(11).Enabled = Not LastRedo
End Sub

Private Sub OpColor_Click(Index As Integer)
Dim sTmp As String
ColorIndex = Index

sTmp = Right("000000" & Hex(OpColor(Index).BackColor), 6)
ScrCol(0).Value = Int("&H" & Right$(sTmp, 2))
ScrCol(1).Value = Int("&H" & Mid$(sTmp, 3, 2))
ScrCol(2).Value = Int("&H" & Left$(sTmp, 2))
End Sub

Private Sub OpColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim sTmp As String
sTmp = Right("000000" & Hex(OpColor(Index).BackColor), 6)
LblColor.Caption = "Hex:" & sTmp & vbCrLf & " Red:" & Int("&H" & Right$(sTmp, 2)) & _
" - Green:" & Int("&H" & Mid$(sTmp, 3, 2)) & " - Blue:" & Int("&H" & Left$(sTmp, 2))
End Sub


Private Sub PicProperty1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
LblColor.Caption = ""
End Sub


Private Sub ScrCol_Change(Index As Integer)
Dim tColor As Long

TxtColor(Index).Text = ScrCol(Index).Value

tColor = RGB(ScrCol(0).Value, ScrCol(1).Value, ScrCol(2).Value)

OpColor(ColorIndex).BackColor = tColor
bFillColor = OpColor(0).BackColor
bBorderColor = OpColor(1).BackColor
bBackColor = OpColor(2).BackColor

If doNothing = True Then Exit Sub

Select Case ColorIndex
    Case 0
    If ObjDraw1.CurrentObject > -1 Then
    ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
    End If
    Case 1
    If ObjDraw1.CurrentObject > -1 Then
    ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
    End If
    Case 2
    ObjDraw1.BackColor = bBackColor
End Select

End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If doNothing = True Then Exit Sub
ObjDraw1.ModifyObject , , , , CSng(Slider1.Value)
End Sub

Private Sub Slider1_Scroll()
Label1(0).Caption = "Rotation: " & Slider1.Value & "°"
End Sub


Private Sub SmnuEdit_Click(Index As Integer)
Select Case Index
    Case 0
    ObjDraw1.Undo
    Case 1
    ObjDraw1.Redo
    Case 2
    'Separator
    Case 3
    ObjDraw1.CopyObject
    ObjDraw1.DeleteObj
    Case 4
    ObjDraw1.CopyObject
    Case 5
    ObjDraw1.PasteObject
    Case 6
    'Separator
    Case 7
    ObjDraw1.DeleteObj
    Case 8
    'separator
    Case 9
    ObjDraw1.SelectAllObjects
    Case 10
    'separator
    Case 11
    ObjDraw1.GroupObjects
    Case 12
    ObjDraw1.UnGroupObjects
End Select
End Sub

Private Sub SmnuFile_Click(Index As Integer)
Select Case Index
    Case 0
    ObjDraw1.CanvasHeight = 480
    ObjDraw1.CanvasWidth = 640
    ObjDraw1.NewProject
    Modified = False
    Case 1
    With cDialog
    .DialogTitle = "Open Project"
    .Filter = "Object Draw Project File (*.ojp)|*.ojp"
    .FileName = ""
    .ShowOpen
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 And FileExist(.FileName) = True Then
        ObjDraw1.OpenProjects .FileName
        bFillColor = ObjDraw1.BackColor
        End If
    End With
    Modified = False
    Case 2
    With cDialog
    .DialogTitle = "Save Project"
    .Filter = "Object Draw Project File (*.ojp)|*.ojp"
    .FileName = ""
    .ShowSave
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 Then
        ObjDraw1.SaveProjects .FileName
        End If
    End With
    Modified = False
    Case 3
    With cDialog
    .DialogTitle = "Export As BitMap"
    .Filter = "Bitmap Image File (*.bmp)|*.bmp"
    .FileName = ""
    .ShowSave
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 Then
        ObjDraw1.Export2BMP .FileName
        End If
    End With
    Case 4
    'Separator
    Case 5
    ObjDraw1.UnSelectAll
    Printer.PaintPicture ObjDraw1.Image, 0, 0
    Printer.EndDoc
    Case 6
    'Separator
    Case 7
    Unload Me
End Select
End Sub


Private Sub SmnuOptions_Click(Index As Integer)
Select Case Index
    Case 0
    Form2.Show vbModal, Me
End Select
End Sub

Private Sub SmnuZoom_Click(Index As Integer)
Dim n As Integer
Select Case Index
    Case 0
    ObjDraw1.ZoomFactor = 0.1
    Case 1
    ObjDraw1.ZoomFactor = 0.25
    Case 2
    ObjDraw1.ZoomFactor = 0.5
    Case 3
    ObjDraw1.ZoomFactor = 1
    Case 4
    ObjDraw1.ZoomFactor = 1.5
    Case 5
    ObjDraw1.ZoomFactor = 2
    Case 6
    ObjDraw1.ZoomFactor = 4
End Select

For n = 0 To 6
SmnuZoom(n).Checked = False
Next n
SmnuZoom(Index).Checked = True
mnuZoom.Caption = "Zoom (" & Round(ObjDraw1.ZoomFactor * 100) & "%)"
StatusBar1.Panels(1).Text = ""
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If doNothing = True Then Exit Sub
Select Case Button.Index
    Case 1
    ObjDraw1.CanvasHeight = 480
    ObjDraw1.CanvasWidth = 640
    ObjDraw1.NewProject
    Modified = False
    Case 2
    With cDialog
    .DialogTitle = "Open Project"
    .Filter = "Object Draw Project File (*.ojp)|*.ojp"
    .FileName = ""
    .ShowOpen
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 And FileExist(.FileName) = True Then
        ObjDraw1.OpenProjects .FileName
        bBackColor = ObjDraw1.BackColor
        OpColor(2).BackColor = bBackColor
        End If
    End With
    Modified = False
    Case 3
    With cDialog
    .DialogTitle = "Save Project"
    .Filter = "Object Draw Project File (*.ojp)|*.ojp"
    .FileName = ""
    .ShowSave
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 Then
        ObjDraw1.SaveProjects .FileName
        End If
    End With
    Modified = False
    Case 4
    With cDialog
    .DialogTitle = "Export As BitMap"
    .Filter = "Bitmap Image File (*.bmp)|*.bmp"
    .FileName = ""
    .ShowSave
    .FileName = Trim(.FileName)
        If Len(.FileName) > 0 Then
        ObjDraw1.Export2BMP .FileName
        End If
    End With
    Case 5
    'separator
    Case 6
    ObjDraw1.CopyObject
    ObjDraw1.DeleteObj
    Case 7
    ObjDraw1.CopyObject
    Case 8
    ObjDraw1.PasteObject
    Case 9
    'separator
    Case 10
    ObjDraw1.Undo
    Case 11
    ObjDraw1.Redo
    Case 12
    'separator
    Case 13
    ObjDraw1.DeleteObj
    Case 14
    'separator
    Case 15
    mTxtAlign = vbLeftJustify
    ObjDraw1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign
    Case 16
    mTxtAlign = vbCenter
    ObjDraw1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign
    Case 17
    mTxtAlign = vbRightJustify
    ObjDraw1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign
    Case 18
    'separator
    Case 19
    mBold = Toolbar1.Buttons(19).Value
    ObjDraw1.ModifyObject , , , , , , , , , , , , Abs(mBold)
    Case 20
    mItalic = Toolbar1.Buttons(20).Value
    ObjDraw1.ModifyObject , , , , , , , , , , , , , Abs(mItalic)
    Case 21
    mUnderline = Toolbar1.Buttons(21).Value
    ObjDraw1.ModifyObject , , , , , , , , , , , , , , Abs(mUnderline)
    Case 22
    mStrikethru = Toolbar1.Buttons(22).Value
    ObjDraw1.ModifyObject , , , , , , , , , , , , , , , Abs(mStrikethru)
    Case 23
    'separator
End Select
UpdToolBar
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
    ObjDraw1.SelectAllObjects
    Case 2
    ObjDraw1.UnSelectAll
    Case 3
    'separator
    Case 4
    ObjDraw1.AlignSelectedObjects mLeft
    Case 5
    ObjDraw1.AlignSelectedObjects mCenter_V
    Case 6
    ObjDraw1.AlignSelectedObjects mRight
    Case 7
    ObjDraw1.AlignSelectedObjects mTop
    Case 8
    ObjDraw1.AlignSelectedObjects mCenter_H
    Case 9
    ObjDraw1.AlignSelectedObjects mBottom
    Case 10
    ObjDraw1.AlignSelectedObjects mCenter_V_H
    Case 11
    'separator
    Case 12
    ObjDraw1.SetObjectOrder ObjDraw1.CurrentObject, BringToFront
    Case 13
    ObjDraw1.SetObjectOrder ObjDraw1.CurrentObject, SendToBack
    Case 14
    ObjDraw1.SetObjectOrder ObjDraw1.CurrentObject, BringFoward
    Case 15
    ObjDraw1.SetObjectOrder ObjDraw1.CurrentObject, SendBackward
    Case 16
    'separator
    Case 17
    ObjDraw1.GroupObjects
    Case 18
    ObjDraw1.UnGroupObjects
End Select
UpdToolBar
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
    ObjDraw1.ZoomFactor = 1
    Case 2
    ObjDraw1.ZoomFactor = ObjDraw1.ZoomFactor - 0.1
    Case 3
    ObjDraw1.ZoomFactor = ObjDraw1.ZoomFactor + 0.1
End Select
mnuZoom.Caption = "Zoom (" & Round(ObjDraw1.ZoomFactor * 100) & "%)"
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tFillColor As Long
Dim tbSize As Integer

tFillColor = bFillColor
tbSize = VScroll1.Value

If tbSize = 0 Then tbSize = 1
Select Case Button.Index
    Case 1
    ObjDraw1.UseSelector
    Case 2
    ObjDraw1.AddObject mline, , , , , , , , bBorderColor, tbSize
    Case 3
    ObjDraw1.AddObject mArc, , , , , CSng(Slider1.Value), , , bBorderColor, tbSize
    Case 4
    ObjDraw1.AddObject mRectangle, , , , , CSng(Slider1.Value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize
    Case 5
    ObjDraw1.AddObject mRoundRectangle, , , , , CSng(Slider1.Value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize, , , , , , , , , , VScroll3.Value
    Case 6
    ObjDraw1.AddObject mEllipse, , , , , CSng(Slider1.Value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize
    StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Circle"
    Case 7
    ObjDraw1.AddObject mPolygon, , , , , CSng(Slider1.Value), tFillColor, CboFill.ListIndex, bBorderColor, VScroll1.Value, , , , , , , , , , bPtsQty
    StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Polygon"
    Case 8
    ObjDraw1.AddObject mStar, , , , , CSng(Slider1.Value), tFillColor, CboFill.ListIndex, bBorderColor, VScroll1.Value, , , , , , , , , , bPtsQty
    StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Polygon"
    Case 9
    ObjDraw1.AddObject mText, , , , , CSng(Slider1.Value), bFillColor, CboFill.ListIndex, , , , CboFontName.Text, _
    CboFontSize.Text, mBold, mItalic, mUnderline, mStrikethru, , mTxtAlign
    Case 10
    With cDialog
    .DialogTitle = "Import Image File"
    .Filter = "All Picture files|*.jpg;*.bmp;*.gif;*.ico;*.cur;*.dib;*.wmf;*.emf"
    .FileName = ""
    .ShowOpen
        If FileExist(.FileName) = True Then
        PicLoad.Picture = LoadPicture(.FileName)
        DoEvents
        ObjDraw1.AddObject mImage, 1, 1, , , , , CboFill.ListIndex, , , PicLoad.Picture
        End If
    End With
End Select
End Sub






Private Sub VScroll1_Change()
If doNothing = True Then Exit Sub
TxtBorder.Text = VScroll1.Value
If ObjDraw1.CurrentObject > -1 Then
ObjDraw1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor, VScroll1.Value
End If
ObjDraw1.SetFocus
End Sub



Private Sub UpdToolBar()
If ObjDraw1.CurrentObject > -1 Then
Toolbar1.Buttons(6).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Toolbar1.Buttons(13).Enabled = True

Toolbar2.Buttons(12).Enabled = True
Toolbar2.Buttons(13).Enabled = True
Toolbar2.Buttons(14).Enabled = True
Toolbar2.Buttons(15).Enabled = True
Toolbar2.Buttons(17).Enabled = True
Toolbar2.Buttons(18).Enabled = True
Else
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(13).Enabled = False

Toolbar2.Buttons(12).Enabled = False
Toolbar2.Buttons(13).Enabled = False
Toolbar2.Buttons(14).Enabled = False
Toolbar2.Buttons(15).Enabled = False
Toolbar2.Buttons(17).Enabled = False
Toolbar2.Buttons(18).Enabled = False
End If

If ObjDraw1.ObjectType = mText Then
Toolbar1.Buttons(15).Enabled = True
Toolbar1.Buttons(16).Enabled = True
Toolbar1.Buttons(17).Enabled = True
Toolbar1.Buttons(19).Enabled = True
Toolbar1.Buttons(20).Enabled = True
Toolbar1.Buttons(21).Enabled = True
Toolbar1.Buttons(22).Enabled = True
CboFontName.Enabled = True
CboFontSize.Enabled = True
Else
Toolbar1.Buttons(15).Enabled = False
Toolbar1.Buttons(16).Enabled = False
Toolbar1.Buttons(17).Enabled = False
Toolbar1.Buttons(19).Enabled = False
Toolbar1.Buttons(20).Enabled = False
Toolbar1.Buttons(21).Enabled = False
Toolbar1.Buttons(22).Enabled = False
CboFontName.Enabled = False
CboFontSize.Enabled = False
End If

Toolbar1.Buttons(8).Enabled = ObjDraw1.ObjectInClipboard
Toolbar2.Buttons(1).Enabled = ObjDraw1.ObjectQty
Toolbar2.Buttons(2).Enabled = ObjDraw1.SelectionQty
If ObjDraw1.SelectionQty > 1 Then
Toolbar2.Buttons(4).Enabled = True
Toolbar2.Buttons(5).Enabled = True
Toolbar2.Buttons(6).Enabled = True
Toolbar2.Buttons(7).Enabled = True
Toolbar2.Buttons(8).Enabled = True
Toolbar2.Buttons(9).Enabled = True
Toolbar2.Buttons(10).Enabled = True
Else
Toolbar2.Buttons(4).Enabled = False
Toolbar2.Buttons(5).Enabled = False
Toolbar2.Buttons(6).Enabled = False
Toolbar2.Buttons(7).Enabled = False
Toolbar2.Buttons(8).Enabled = False
Toolbar2.Buttons(9).Enabled = False
Toolbar2.Buttons(10).Enabled = False
End If
End Sub

Private Sub VScroll2_Change()
If doNothing = True Then Exit Sub
TxtPoint.Text = VScroll2.Value
bPtsQty = VScroll2.Value
If ObjDraw1.CurrentObject > -1 And ObjDraw1.ObjectType = mPolygon Or ObjDraw1.ObjectType = mStar Then
ObjDraw1.ModifyObject , , , , , , , , , , , , , , , , , , VScroll2.Value
End If
ObjDraw1.SetFocus
End Sub


Private Sub VScroll3_Change()
If doNothing = True Then Exit Sub
TxtRound.Text = VScroll3.Value
If ObjDraw1.CurrentObject > -1 And ObjDraw1.ObjectType = mRoundRectangle Then
ObjDraw1.ModifyObject , , , , , , , , , , , , , , , , , , VScroll3.Value
End If
ObjDraw1.SetFocus
End Sub


