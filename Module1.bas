Attribute VB_Name = "Module1"
Sub Main()
    
    App.Title = appTitle
    
    frmSplash.Show
    frmSplash.Refresh
    
    Set frmMain = New frmMain
    Load frmMain
    Unload frmSplash


    frmMain.Show
End Sub
