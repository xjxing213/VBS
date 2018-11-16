const TriggerTypeTime = 1
const ActionTypeExec = 0   

'********************************************************
Set service = CreateObject("Schedule.Service")
call service.Connect()

'********************************************************
Dim rootFolder
Set rootFolder = service.GetFolder("\")

Dim taskDefinition
Set taskDefinition = service.NewTask(0) 

'********************************************************
Dim regInfo
Set regInfo = taskDefinition.RegistrationInfo
regInfo.Description = "Start Disk Defragmenter"
regInfo.Author = "Administrator"

Dim settings
Set settings = taskDefinition.Settings
settings.Enabled = True
settings.StartWhenAvailable = True
settings.Hidden = False

'********************************************************
Dim triggers
Set triggers = taskDefinition.Triggers

Dim trigger
Set trigger = triggers.Create(TriggerTypeTime)

Dim startTime, endTime

Dim time
time = DateAdd("s", 30, Now)  
startTime = XmlTime(time)

time = DateAdd("n", 15, Now) 
endTime = XmlTime(time)

WScript.Echo "startTime :" & startTime
WScript.Echo "endTime :" & endTime

trigger.StartBoundary = startTime
trigger.EndBoundary = endTime
trigger.ExecutionTimeLimit = "PT60M"    
trigger.Id = "TimeTriggerId"
trigger.Enabled = True

'***********************************************************

Dim Action
Set Action = taskDefinition.Actions.Create( ActionTypeExec )
Action.Path = "C:\Windows\System32\dfrgui.exe"

WScript.Echo "Task definition created ... submitting task..."

'***********************************************************

call rootFolder.RegisterTaskDefinition( _
    "Test TimeTrigger", taskDefinition, 6, , , 3)

WScript.Echo "Task submitted."

Function XmlTime(t)
    Dim cSecond, cMinute, CHour, cDay, cMonth, cYear
    Dim tTime, tDate

    cSecond = "0" & Second(t)
    cMinute = "0" & Minute(t)
    cHour = "0" & Hour(t)
    cDay = "0" & Day(t)
    cMonth = "0" & Month(t)
    cYear = Year(t)

    tTime = Right(cHour, 2) & ":" & Right(cMinute, 2) & _
        ":" & Right(cSecond, 2)
    tDate = cYear & "-" & Right(cMonth, 2) & "-" & Right(cDay, 2)
    XmlTime = tDate & "T" & tTime 
End Function