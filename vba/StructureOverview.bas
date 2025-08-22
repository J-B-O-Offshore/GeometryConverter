Attribute VB_Name = "StructureOverview"
Sub show_GEOMETRY_section()
    ShowOnlySelectedColumns "D:BE", "D:P"
End Sub

Sub show_ADDED_MASSES_section()
    ShowOnlySelectedColumns "D:BE", "Q:z"
End Sub

Sub show_MISC_section()
    ShowOnlySelectedColumns "D:BE", "AB:AJ"
End Sub

Sub plot_Assambly_Overview()
    RunPythonWrapper "plot", "plot_Assambly_Overview"
End Sub

