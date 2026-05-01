# Chapter 5: Advanced Control Types

As your Add-in grows, you need more than just standard buttons. Advanced controls help you save space, manage settings, and capture user input directly from the Ribbon.

---

## 1. Menus & SplitButtons (Grouping Actions)
Use a `menu` to group multiple related macros under one icon. A `splitButton` is similar but allows the top icon to be a clickable button itself.

**XML:**
```xml
<!-- A simple dropdown menu -->
<menu id="exportMenu" label="Export Data" imageMso="ExportExcel" size="large">
    <button id="btnPdf" label="As PDF" onAction="Ribbon_Router" tag="pdf" />
    <button id="btnCsv" label="As CSV" onAction="Ribbon_Router" tag="csv" />
</menu>

<!-- A button with a dropdown attached -->
<splitButton id="split1" size="large">
    <button id="mainBtn" label="Fast Run" onAction="MyMacro" imageMso="Play" />
    <menu id="subOptions">
        <button id="opt1" label="Run with Logs" onAction="MyMacro" />
    </menu>
</splitButton>
```

---

## 2. Checkboxes & ToggleButtons (State Management)
These controls store an "On" or "Off" state. They use the `getPressed` callback to show their current status.

**XML:**
```xml
<toggleButton id="tglDraft" label="Draft Mode" onAction="OnToggle" getPressed="GetTglStatus" imageMso="ReviewUpdate" />
<checkBox id="chkAuto" label="Auto-Save" onAction="OnCheck" getPressed="GetChkStatus" />
```

**VBA:**
```vba
' VBA pushes the True/False state to the control
Public Sub GetTglStatus(control As IRibbonControl, ByRef returnedVal)
    returnedVal = MyGlobalStateVariable 
End Sub

' When clicked, the 'pressed' parameter tells you the new state
Public Sub OnToggle(control As IRibbonControl, pressed As Boolean)
    MyGlobalStateVariable = pressed
    MyRibbonUI.Invalidate ' Refresh to show the new state visually
End Sub
```

---

## 3. EditBoxes & ComboBoxes (User Input)
If you need a user to type a value (like a percentage or a name) without leaving the Ribbon, use an `editBox`.

**XML:**
```xml
<editBox id="txtThreshold" 
         label="Limit:" 
         getText="GetDefaultText" 
         onChange="OnTextChange" />
```

**VBA:**
```vba
' Captures the text typed by the user
Public Sub OnTextChange(control As IRibbonControl, text As String)
    MsgBox "You entered: " & text
    ' Update your logic or variables here
End Sub

' Provides the starting text for the box
Public Sub GetDefaultText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "100"
End Sub
```

---

## 4. Visual Aesthetics (Separators & Labels)
Good UI design requires spacing. Use these tags to organize your groups.

*   **Separator:** A vertical line that divides buttons into clusters.
*   **LabelControl:** Static text used for headers or instructions.

**XML:**
```xml
<group id="grpLogic" label="Analysis Tools">
    <button id="btnStart" label="Start" imageMso="Play" />
    
    <separator id="sep1" />
    
    <labelControl id="lblNote" label="Internal Use Only" />
    <button id="btnEnd" label="Stop" imageMso="Stop" />
</group>
```

---

## Summary of Callbacks for Advanced Types

| Control | XML Callback | VBA Event |
| :--- | :--- | :--- |
| **Menu/Button** | `onAction` | Triggered when clicked. |
| **Toggle/Check** | `getPressed` | Sets visual "On/Off" state. |
| **EditBox** | `onChange` | Captures typed text string. |
| **All Types** | `getEnabled` | Controls if the user can interact with it. |
