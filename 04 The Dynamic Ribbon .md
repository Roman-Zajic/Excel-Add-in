# Chapter 4: The Dynamic Ribbon (Refreshing the UI)

By default, the Ribbon is "set in stone" once it loads. To change a button’s state (e.g., graying it out or hiding it) while Excel is open, you must force the Ribbon to refresh using the **Invalidate** command.

---

## 1. The Refresh Command (`Invalidate`)
To update the Ribbon, you call `Invalidate`. This forces Excel to re-read your VBA logic and redraw the UI.

**VBA:**
```vba
' This sub forces the entire Ribbon to refresh its state
Public Sub RefreshRibbon()
    If Not MyRibbonUI Is Nothing Then
        MyRibbonUI.Invalidate 
    End If
End Sub
```
*Note: `MyRibbonUI` is the variable we captured in Chapter 3.*

---

## 2. Enabling and Disabling Buttons
If you want a button to be clickable only under certain conditions (like when a specific cell is filled), use `getEnabled`.

**XML:**
```xml
<button id="btnProcess" label="Process" onAction="MyMacro" getEnabled="IsButtonEnabled" />
```

**VBA:**
```vba
' Excel calls this every time the Ribbon is invalidated
Public Sub IsButtonEnabled(control As IRibbonControl, ByRef returnedVal)
    If Range("A1").Value <> "" Then
        returnedVal = True  ' Button is clickable
    Else
        returnedVal = False ' Button is grayed out
    End If
End Sub
```

---

## 3. Hiding and Showing Elements
You can hide entire groups or buttons using `getVisible`. This is useful for "Admin-only" tools.

**XML:**
```xml
<group id="AdminTools" label="Admin" getVisible="ShowAdminGroup">
    <button id="btnSecret" label="Clear Logs" onAction="MyMacro" />
</group>
```

**VBA:**
```vba
Public Sub ShowAdminGroup(control As IRibbonControl, ByRef returnedVal)
    ' Hide the group unless the user is "Manager"
    returnedVal = (Application.UserName = "Manager")
End Sub
```

---

## 4. The 3-Step Dynamic Workflow

To understand how to make the Ribbon react to your actions, let’s build a **"Safety Lock"** system: One button "Unlocks" a second button.

### Step 1: Define the XML
We define two buttons. The second one uses `getEnabled` instead of a static `enabled` attribute.

```xml
<group id="SecurityGroup" label="Safety System">
    <!-- This button toggles the lock -->
    <button id="btnToggleLock" label="Toggle Lock" onAction="OnToggleLock" imageMso="Lock" />
    
    <!-- This button is controlled by VBA via 'getEnabled' -->
    <button id="btnAction" label="Dangerous Action" onAction="OnRunAction" getEnabled="GetActionStatus" imageMso="Delete" />
</group>
```

### Step 2: Set up the VBA State and Callback
In your VBA module, create a variable to track the "Lock" state and a sub to report that state back to the Ribbon.

```vba
Public IsLocked As Boolean ' Tracks if the tool is locked

' Ribbon calls this to decide if btnAction should be grayed out
Public Sub GetActionStatus(control As IRibbonControl, ByRef returnedVal)
    If IsLocked Then
        returnedVal = False ' Button is grayed out
    Else
        returnedVal = True  ' Button is clickable
    End If
End Sub
```

### Step 3: Trigger the Refresh (The Toggle)
Now, create the macro for the first button. It must change the `IsLocked` variable and then tell the Ribbon to refresh itself.

```vba
' Called when you click "Toggle Lock"
Public Sub OnToggleLock(control As IRibbonControl)
    ' 1. Change the logic state
    IsLocked = Not IsLocked 
    
    ' 2. Tell the user what happened
    MsgBox "System is now " & IIf(IsLocked, "LOCKED", "UNLOCKED")
    
    ' 3. FORCE the Ribbon to refresh and re-run "GetActionStatus"
    If Not MyRibbonUI Is Nothing Then
        MyRibbonUI.Invalidate 
    End If
End Sub

' The actual tool
Public Sub OnRunAction(control As IRibbonControl)
    MsgBox "Dangerous action performed!"
End Sub
```

### Why this works:
1.  When Excel starts, it runs `GetActionStatus` and sees `IsLocked` is False (default).
2.  When you click **Toggle Lock**, the VBA code flips the variable to True.
3.  The `MyRibbonUI.Invalidate` command acts like a "Refresh" button. 
4.  Excel immediately looks at the XML, sees `getEnabled="GetActionStatus"`, and runs that VBA sub again. 
5.  Since the variable is now True, the "Dangerous Action" button instantly turns gray.
*   **Static:** Attributes like `label="Run"` never change.
*   **Dynamic:** Attributes like `getLabel="SubName"` call VBA every time the Ribbon refreshes.
*   **Invalidate:** The "Trigger" that tells Excel to run all your "Get" macros right now.
