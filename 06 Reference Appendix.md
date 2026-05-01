# Chapter 6: Reference Appendix

This appendix serves as a quick-lookup guide for the most common Ribbon controls and their corresponding VBA signatures.

---

## 1. Gallery of imageMso (Icons)
The `imageMso` attribute allows you to use thousands of high-quality icons built into Microsoft Office. You do not need to include any image files in your project to use these.

### Where to find Icon IDs:
*   **Official Microsoft Download:** [Office 2010 Icon Gallery](https://www.microsoft.com/en-us/download/details.aspx?id=21103) (An Excel file containing all icons).

---

## 2. XML Control Cheat Sheet
Copy and paste these snippets directly into your `customUI14.xml`.

### Standard Button
```xml
<button id="btn1" label="Run Tool" size="large" imageMso="HappyFace" onAction="MyMacro" />
```
**VBA:** `Sub MyMacro(control As IRibbonControl)`

### Toggle Button (Stay Pressed)
```xml
<toggleButton id="tgl1" label="Mode On/Off" size="large" imageMso="AdpPrimaryKey" onAction="OnToggle" getPressed="GetStatus" />
```
**VBA:** `Sub OnToggle(control As IRibbonControl, pressed As Boolean)`

### Checkbox
```xml
<checkBox id="chk1" label="Enable Auto-Run" onAction="OnCheck" getPressed="GetStatus" />
```
**VBA:** `Sub OnCheck(control As IRibbonControl, pressed As Boolean)`

### Menu (Dropdown)
```xml
<menu id="menu1" label="Options" imageMso="Filter" size="large">
    <button id="opt1" label="Sub Action A" onAction="MyMacro" />
    <button id="opt2" label="Sub Action B" onAction="MyMacro" />
</menu>
```

### Split Button
```xml
<splitButton id="split1" size="large">
    <button id="mainBtn" label="Primary" onAction="MainMacro" imageMso="Play" />
    <menu id="subOptions">
        <button id="optA" label="Secondary" onAction="SubMacro" />
    </menu>
</splitButton>
```

### Edit Box (Text Input)
```xml
<editBox id="txt1" label="Enter Name:" onChange="OnTextChange" getText="GetDefaultText" />
```
**VBA:** `Sub OnTextChange(control As IRibbonControl, text As String)`

### Combo Box (Dropdown + Input)
```xml
<comboBox id="combo1" label="Pick Choice:" onChange="OnComboChange">
    <item id="i1" label="Choice 1" />
    <item id="i2" label="Choice 2" />
</comboBox>
```

---

## 3. Organizational Tags
These tags do not trigger macros but control the visual spacing of your Ribbon.

| Tag | Purpose |
| :--- | :--- |
| `<separator id="s1" />` | A vertical line to separate buttons. |
| `<labelControl id="l1" label="Status: Active" />` | Static text that doesn't click. |
| `<group id="g1" label="My Tools">` | The container that holds all controls. |
| `insertAfterMso="GroupEditing"` | Used inside a `group` tag to place your group after a specific built-in Excel group. |

---

## 4. Full Callback Table
Use this as a reference for your VBA parameter lists.

| Callback | VBA Signature |
| :--- | :--- |
| **onAction** | `(control As IRibbonControl)` |
| **onAction (Toggle)** | `(control As IRibbonControl, pressed As Boolean)` |
| **onChange (Edit)** | `(control As IRibbonControl, text As String)` |
| **getLabel** | `(control As IRibbonControl, ByRef returnedVal)` |
| **getEnabled** | `(control As IRibbonControl, ByRef returnedVal)` |
| **getVisible** | `(control As IRibbonControl, ByRef returnedVal)` |
| **getPressed** | `(control As IRibbonControl, ByRef returnedVal)` |
| **onLoad** | `(ribbon As IRibbonUI)` |
