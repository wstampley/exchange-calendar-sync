# exchange-calendar-sync
As a consultant, there are almost always two calendars that I need to clone everything from one to another. This is VBA magic to do just that.

# Usage Instructions

Import the EmailSynchronizer.cls into your class modules for outlook.

Add the following to the 'ThisOutlookSession' module:

```
Option Explicit

Dim EmailSyncInstance As EmailSynchronizer

Private Sub Application_Startup()

    Set EmailSyncInstance = New EmailSynchronizer
    EmailSyncInstance.SetupCalendarSync "will.stampley@client-email.com", "will.stampley@consultant-email.com"

End Sub
```

If you already have an Application_Startup Sub, you will need to just add the contents to the existing one.
