Attribute VB_Name = "DateExpiration"
Public Sub SetExpiryTime()
  Dim Sel As Outlook.Selection
  Dim obj As Object
  Dim Interval As Long
  Dim ExpiryTime As Date
  Dim text$
  Dim frm As DateSelectionBox

    Set frm = New DateSelectionBox

  If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
    Set obj = Application.ActiveInspector.CurrentItem

 Else
    Set Sel = Application.ActiveExplorer.Selection
    If Sel.Count = 0 Then
      Exit Sub
    Else
      Set obj = Sel(1)
    End If
  End If

  Select Case True
  Case (TypeOf obj Is Outlook.mailitem), _
    (TypeOf obj Is Outlook.MeetingItem), _
    (TypeOf obj Is Outlook.PostItem)

    ExpiryTime = obj.ExpiryTime
  End Select

  If ExpiryTime = #1/1/4501# Then
    text = "No Date set"
  Else
    text = ExpiryTime
  End If

  'Text = "Aktuelles Ablaufdatum: " & Text & vbCrLf & vbCrLf
  'Text = Text & "In wieviel Wochen soll die Auswahl ablaufen?"
  'Text = InputBox(Text, , "8")
  
  frm.TextBoxCurrentExpDate.Value = text
  frm.MonthView1.Value = Date
  frm.Show
  

  If frm.Applied Then
    ExpiryTime = frm.AppliedDate

    If Not Sel Is Nothing Then
      For Each obj In Sel

        Select Case True
        Case (TypeOf obj Is Outlook.mailitem), _
          (TypeOf obj Is Outlook.MeetingItem), _
          (TypeOf obj Is Outlook.PostItem)

          obj.ExpiryTime = ExpiryTime
          obj.Save
        End Select
      Next

    Else
      Select Case True
      Case (TypeOf obj Is Outlook.mailitem), _
        (TypeOf obj Is Outlook.MeetingItem), _
        (TypeOf obj Is Outlook.PostItem)

        obj.ExpiryTime = ExpiryTime
        obj.Save
      End Select
    End If
  End If
  
  Unload frm
End Sub

