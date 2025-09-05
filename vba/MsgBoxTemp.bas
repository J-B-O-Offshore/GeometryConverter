Attribute VB_Name = "MsgBoxTemp"

    Function ShowMessageBox() As Integer
        ShowMessageBox = MsgBox("The path '\\\\HH-DC02.jbo.local\daten\Offshore\34_Geometry_Databases\TP.db' does not lead to a valid SQLite database. Try reloading the database.", 64, "Message")
    End Function
    
