VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgGisFormat 
   Caption         =   "Export GIS"
   ClientHeight    =   2040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4980
   OleObjectBlob   =   "dlgGisFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgGisFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Export format prompt (KML / KMZ / Shapefile), like the Python tool's Export GIS dialog.
Public result As String

Private Sub UserForm_Initialize()
    cboFormat.AddItem "KML + Shapefile"
    cboFormat.AddItem "KMZ + Shapefile"
    cboFormat.AddItem "KML"
    cboFormat.AddItem "KMZ"
    cboFormat.AddItem "Shapefile"
    cboFormat.ListIndex = 0
End Sub

Private Sub btnExport_Click()
    result = cboFormat.text
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    result = ""
    Me.Hide
End Sub

