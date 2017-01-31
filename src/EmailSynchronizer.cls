VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailSynchronizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MonitoredCalendarItems As Outlook.Items
Attribute MonitoredCalendarItems.VB_VarHelpID = -1
Dim WithEvents DeletedItems As Outlook.Items
Attribute DeletedItems.VB_VarHelpID = -1
Dim TargetCalendar As Outlook.Folder

Private Const MeetingPrefix As String = "Client: "
Private Const MeetingCategory As String = "Client Sync"
Private Const GUIDPrefix As String = "CSGUID:"
Private Const GUIDLength As Integer = 36
Private Const GUIDPrefixLength As Integer = 7
Private Const TotalGUIDSize As Integer = GUIDLength + GUIDPrefixLength + 2

Public Sub SetupCalendarSync(ByVal monitoredAccount As String, ByVal targetAccount As String)

    If Not MonitoredCalendarItems Is Nothing And Not DeletedItems Is Nothing And Not TargetCalendar Is Nothing Then
        'Already setup, do nothing
        Exit Sub
    End If

    Dim NS As Outlook.NameSpace
    Set NS = Application.GetNamespace("MAPI")

    Dim monitoredCalendarPath As String
    Dim monitoredDeletedItemsPath As String
    monitoredCalendarPath = monitoredAccount + "\calendar"
    monitoredDeletedItemsPath = monitoredAccount + "\Deleted Items"

    Dim targetCalendarPath As String
    targetCalendarPath = targetAccount + "\calendar"

    ' calendar to watch for new items
    Dim monitoredCalendar As Outlook.Folder
    Set monitoredCalendar = GetFolderFromPath(monitoredCalendarPath)
    Set MonitoredCalendarItems = monitoredCalendar.Items
    Set DeletedItems = GetFolderFromPath(monitoredDeletedItemsPath).Items

    ' calendar moving copy to
    Set TargetCalendar = GetFolderFromPath(targetCalendarPath)

    Set NS = Nothing
End Sub

Public Sub PerformOneTimeSync()

    If MonitoredCalendarItems Is Nothing Then
        Exit Sub
    End If

    If TargetCalendar Is Nothing Then
        Exit Sub
    End If

    Dim monitoredCalendarAppointment As AppointmentItem
    For Each monitoredCalendarAppointment In MonitoredCalendarItems
        If monitoredCalendarAppointment.Start >= Now() Then
            Dim clonedCalendarAppointment As AppointmentItem
            Set clonedCalendarAppointment = FindClonedAppointmentForAppointment(monitoredCalendarAppointment)

            If clonedCalendarAppointment Is Nothing Then
                CreateCloneAppointmentForAppointment monitoredCalendarAppointment
            End If
        End If
    Next

End Sub

Private Sub MonitoredCalendarItems_ItemAdd(ByVal monitoredItem As Object)

    On Error Resume Next

    'Only copy Busy Events (not Tentative, Free, etc.)
    If monitoredItem.BusyStatus <> olBusy Then
        Exit Sub
    End If

    'Do not copy All Day events
    If monitoredItem.AllDayEvent = True Then
        Exit Sub
    End If

    'Otherwise, GUID It, and copy it
    monitoredItem.Body = monitoredItem.Body & vbNewLine & "[" & GUIDPrefix & CreateGUID & "]"
    monitoredItem.Save

    Dim cAppt As AppointmentItem
    Set cAppt = Application.CreateItem(olAppointmentItem)
    With cAppt
        .Subject = MeetingPrefix & monitoredItem.Subject
        .Start = monitoredItem.Start
        .Duration = monitoredItem.Duration
        .Location = monitoredItem.Location
        .Body = monitoredItem.Body
    End With

    ' set the category after moving to force Exchange's ActiveSync to sync changes
    Dim moveCal As AppointmentItem
    Set moveCal = cAppt.Move(TargetCalendar)
    moveCal.Categories = MeetingCategory
    moveCal.Save

End Sub


Private Sub MonitoredCalendarItems_ItemChange(ByVal monitoredItem As Object)
    Dim cAppt As AppointmentItem
    Dim objAppointment As AppointmentItem

    On Error Resume Next

    Set cAppt = FindClonedAppointmentForAppointment(monitoredItem)

    ' Don't do anything if we can't find the sync'd item
    If cAppt Is Nothing Then
        ' Eventually, this would be nice to create the missing sync'd item
        Exit Sub
    End If

    ' Otherwise, since we found the clone, update all of its properties
    With cAppt
        .Subject = MeetingPrefix & monitoredItem.Subject
        .Start = monitoredItem.Start
        .Duration = monitoredItem.Duration
        .Location = monitoredItem.Location
        .Body = monitoredItem.Body
        .Save
    End With

End Sub

Private Sub DeletedItems_ItemAdd(ByVal DeletedItem As Object)
    On Error Resume Next

    ' only apply to appointments
    If DeletedItem.MessageClass <> "IPM.Appointment" Then
        Exit Sub
    End If

    ' If it was one of our sync'd items somehow, we can ignore it
    If DeletedItem.Categories = MeetingCategory Then
        Exit Sub
    End If

    ' Find the synchronized item that has the same GUID in them and delete it
    Dim cAppt As AppointmentItem
    Set cAppt = FindClonedAppointmentForAppointment(DeletedItem)
    If Not cAppt Is Nothing Then
        cAppt.Delete
    End If

End Sub

Private Function FindClonedAppointmentForAppointment(ByVal monitoredAppointment As AppointmentItem) As AppointmentItem

    Set FindClonedAppointmentForAppointment = Nothing

    Dim monitoredItemGUID As String
    monitoredItemGUID = Right(monitoredAppointment.Body, TotalGUIDSize)

    If Len(monitoredItemGUID) = 0 Then
        ' Don't do anything if there isn't a GUID
        Exit Function
    End If

    If InStr(monitoredItemGUID, GUIDPrefix) = 0 Then
        ' Don't do anything if there isn't a GUID
        Exit Function
    End If

    Set FindClonedAppointmentForAppointment = FindClonedAppointmentForGUID(monitoredItemGUID)

End Function

Private Function FindClonedAppointmentForGUID(ByVal GUID As String) As AppointmentItem

    Set FindClonedAppointmentForGUID = Nothing

    Dim targetCalendarAppointment As AppointmentItem
    For Each targetCalendarAppointment In TargetCalendar.Items
        If InStr(targetCalendarAppointment.Body, GUID) Then
            Set FindClonedAppointmentForGUID = targetCalendarAppointment
            Exit For
        End If
    Next

End Function

Private Sub CreateCloneAppointmentForAppointment(ByVal monitoredAppointment As AppointmentItem)



End Sub

Private Function CreateGUID() As String
    CreateGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, GUIDLength)
End Function


Private Function GetFolderFromPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer

    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If

    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderFromPath = Nothing
                Exit Function
            End If
        Next
    End If

    'Return the oFolder
    Set GetFolderFromPath = oFolder
    Exit Function

GetFolderPath_Error:
        Set GetFolderFromPath = Nothing
        Exit Function
End Function
