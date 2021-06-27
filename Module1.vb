Module Module1

    Sub Main()

        Const TriggerTypeDaily = 2
        Const ActionTypeExec = 0
        Dim service As Object
        service = CreateObject("Schedule.Service")
        Call service.Connect()
        Dim rootFolder
        rootFolder = service.GetFolder("\")
.       Dim taskDefinition
        taskDefinition = service.NewTask(0)


        Dim regInfo
        regInfo = taskDefinition.RegistrationInfo
        regInfo.Description = "Start notepad at 8:00AM daily"
        regInfo.Author = "Administrator"

  
        Dim settings
        settings = taskDefinition.Settings
        settings.Enabled = True
        settings.StartWhenAvailable = True
        settings.Hidden = False

   
        Dim triggers
        triggers = taskDefinition.Triggers

        Dim trigger
        trigger = triggers.Create(TriggerTypeDaily)

     
        Dim startTime, endTime

        Dim time
        startTime = "2019-05-02T12:00:00"  
        endTime = "2022-05-02T13:00:00"

 

        trigger.StartBoundary = startTime
        trigger.EndBoundary = endTime
        trigger.DaysInterval = 1   
        trigger.Id = "DailyTriggerId"
        trigger.Enabled = True

        Dim repetitionPattern
        repetitionPattern = trigger.Repetition
        repetitionPattern.Duration = "PT12H"
        repetitionPattern.Interval = "PT5M"

  
        Dim Action
        Action = taskDefinition.Actions.Create(ActionTypeExec)
        Action.Path = "C:\Windows\System32\notepad.exe"

        

        Call rootFolder.RegisterTaskDefinition(
        "Test Daily Trigger", taskDefinition, 6, , , 3)

        
    End Sub

End Module
