VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Running"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    'hide the title bar if you're working on a windows machine. Otherwise, just display it as you normally would
    Me.Height = Me.Height - 10
   

End Sub
