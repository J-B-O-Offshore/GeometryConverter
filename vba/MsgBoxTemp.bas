Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("Please set a water level in StructureOverview, as you set water_level in config1 to 'auto'. Aborting.", 64, "Message")
    End Function
    
