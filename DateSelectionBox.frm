VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateSelectionBox 
   Caption         =   "Select Expiration Date"
   ClientHeight    =   3718
   ClientLeft      =   33
   ClientTop       =   363
   ClientWidth     =   4862
   OleObjectBlob   =   "DateSelectionBox.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "DateSelectionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Applied As Boolean
Public AppliedDate As Date

Private Sub ButtonApply_Click()
    Applied = True
    AppliedDate = MonthView1.Value
    Me.Hide
End Sub

Private Sub ButtonCancel_Click()
    Applied = False
    
    Me.Hide
End Sub

Private Sub ButtonDelExpDate_Click()
    Applied = True
    AppliedDate = #1/1/4501#
    Me.Hide
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    If DateClicked < Date Then
        MonthView1.Value = Date
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
    ' user clicked the X button
    ' cancel unloading the form, use close button procedure instead
    Cancel = True
    ButtonCancel_Click
  End If
End Sub
