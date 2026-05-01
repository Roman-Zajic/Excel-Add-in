# Chapter 3: The Callback System

In a Ribbon Add-in, the XML is the "Front-end" (what the user sees) and VBA is the "Back-end" (the logic). To connect them, we use **Callbacks**.

---

## 1. The Mandatory Signature
For a button to trigger a VBA macro, the sub **must** accept an `IRibbonControl` object. If this parameter is missing, the button will fail to find the macro.

```vba
' Standard signature for all Ribbon actions
Public Sub MyMacro(control As IRibbonControl)
    MsgBox "Clicked!"
End Sub
```

---

## 2. Using the "Traffic Controller" Pattern (Tag vs ID)
Instead of writing 50 different macros for 50 buttons, you can use a single "Traffic Controller" macro to route actions. 

**XML Setup:**
```xml
<button id="btn_01" tag="run_report" onAction="Ribbon_Router" label="Run Report" />
<button id="btn_02" tag="clear_data" onAction="Ribbon_Router" label="Clear All" />
```

**VBA Router:**
```vba
Public Sub Ribbon_Router(control As IRibbonControl)
    ' We use the "Tag" attribute because it is more flexible than the ID.
    ' IDs must be unique (btn_01, btn_02), but Tags can be grouped logic.
    
    Select Case control.Tag
        Case "run_report": Call Logic_RunReport
        Case "clear_data": Call Logic_ClearData
    End Select
End Sub
```
**Why use Tags?** `IDs` are strictly for the XML to identify elements. `Tags` are your private "data" field. Using Tags allows you to change your XML structure (the ID) without breaking your VBA logic.

---

## 3. The `onLoad` Callback (Gaining Control)
By default, VBA cannot "talk back" to the Ribbon. If you want to force the Ribbon to refresh or change a button's label while Excel is running, you must capture a reference to it during startup.

**XML:**
```xml
<customUI xmlns="..." onLoad="Capture_Ribbon">
```

**VBA:**
```vba
Public MyRibbonUI As IRibbonUI

' Triggered exactly once when the Add-in is loaded
Public Sub Capture_Ribbon(ribbon As IRibbonUI)
    Set MyRibbonUI = ribbon
End Sub
```
**Why do we do this?** Without the `MyRibbonUI` variable, you have no way to trigger a "Refresh" (Invalidate) command later.

---

## 4. "Get" Callbacks (VBA Pushing to Ribbon)
Standard attributes like `label="Save"` are static. **Get** callbacks allow VBA to dynamically "push" values to the Ribbon. This is used to change button text or disable buttons based on conditions.

**XML:**
```xml
<button id="btn1" getLabel="GetDynamicLabel" onAction="Ribbon_Router" />
```

**VBA:**
```vba
' This macro is called by the Ribbon whenever it needs to "know" its label
Public Sub GetDynamicLabel(control As IRibbonControl, ByRef returnedVal)
    If ActiveSheet.Name = "Report" Then
        returnedVal = "Export Report"
    Else
        returnedVal = "General Tool"
    End If
End Sub
```
**Why use this?** You use "Get" callbacks when the UI needs to react to the state of the workbook (e.g., disabling a "Delete" button if the sheet is protected). 

*Note: To force these "Get" macros to re-run, you must call `MyRibbonUI.Invalidate`.*
