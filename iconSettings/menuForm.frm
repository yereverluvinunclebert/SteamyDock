VERSION 5.00
Begin VB.Form menuForm 
   Caption         =   "Rocketdock MenuForm"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMainOpts 
      Caption         =   "Other Options"
      Visible         =   0   'False
      Begin VB.Menu mnuOtherOpts 
         Caption         =   "GDI"
         Begin VB.Menu mnuSubOpts 
            Caption         =   "Don't Use GDI+"
            Index           =   0
         End
         Begin VB.Menu mnuSubOpts 
            Caption         =   "Use GDI+"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSaveOpts 
         Caption         =   "Save As"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using GDI+)"
            Index           =   0
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using zLIB)"
            Index           =   1
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Default Filter"
               Index           =   0
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use No Filters (Fastest)"
               Index           =   1
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Left Filter"
               Index           =   2
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Top Filter"
               Index           =   3
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Average Filter"
               Index           =   4
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Paeth Filter"
               Index           =   5
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adaptive Filtering (Slowest)"
               Index           =   6
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As JPG"
            Index           =   2
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As TGA"
            Index           =   3
            Begin VB.Menu mnuTGA 
               Caption         =   "Compressed"
               Index           =   0
            End
            Begin VB.Menu mnuTGA 
               Caption         =   "Uncompressed"
               Index           =   1
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As GIF"
            Index           =   4
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As BMP (Red Bkg)"
            Index           =   5
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As Rendered Example (GDI+ required)"
            Index           =   7
         End
      End
      Begin VB.Menu mnuPos 
         Caption         =   "Position"
         Enabled         =   0   'False
         Begin VB.Menu mnuPosSub 
            Caption         =   "Centered"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Top Left"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Top Right"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Bottom Left"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Bottom Right"
            Enabled         =   0   'False
            Index           =   4
         End
      End
   End
   Begin VB.Menu rdMapMenu 
      Caption         =   "The Map Menu"
      Visible         =   0   'False
      Begin VB.Menu menuDelete 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu menuAdd 
         Caption         =   "Add Item"
      End
      Begin VB.Menu menuLeft 
         Caption         =   "Move item to the left"
      End
      Begin VB.Menu menuRight 
         Caption         =   "Move item to the right"
      End
   End
   Begin VB.Menu mnupopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About this utility"
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with paypal"
         Index           =   2
      End
      Begin VB.Menu mnuSweets 
         Caption         =   "Donate some sweets/candy with Amazon"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Utility Help"
         Index           =   4
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version"
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "Contact Support"
         End
         Begin VB.Menu mnuMoreIcons 
            Caption         =   "Visit Deviantart to download some more Icons"
         End
         Begin VB.Menu mnuWidgets 
            Caption         =   "See the complementary steampunk widgets"
         End
         Begin VB.Menu mnuFacebook 
            Caption         =   "Chat about Rocketdock functionality on Facebook"
         End
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
   End
   Begin VB.Menu thumbmenu 
      Caption         =   "Thumb Menu"
      Visible         =   0   'False
      Begin VB.Menu menuSmallerIcons 
         Caption         =   "small icons with text"
      End
      Begin VB.Menu menuLargerThumbs 
         Caption         =   "larger icons (no text)"
      End
   End 
End
Attribute VB_Name = "Menuform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnutest_Click()

End Sub
