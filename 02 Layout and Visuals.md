# Chapter 2: Ribbon Layout & Visual Controls

The Ribbon follows a strict hierarchy: **Tab > Group > Control**. This chapter covers how to organize your tools and make them look professional using built-in Office features.

---

## 1. Organizing the Hierarchy
You can create your own tab or inject your groups into existing standard tabs (like the "Home" tab).

**To create a custom tab:**
```xml
<tab id="MyTab" label="DATA TOOLS">
    <group id="Grp1" label="Formatting">
        <!-- Controls go here -->
    </group>
</tab>
```

**To add to a built-in tab (e.g., Home Tab):**
*Use `idMso` instead of `id` to target Microsoft’s internal names.*
```xml
<tab idMso="TabHome">
    <group id="CustomGroupInHome" label="My Add-In Tools" insertAfterMso="GroupEditing">
        <button id="BtnHome1" label="Quick Action" onAction="MyMacro" imageMso="Rocket" />
    </group>
</tab>
```

---

## 2. Control Sizing & Labels
Buttons come in two primary sizes. Organizing them correctly prevents the Ribbon from looking cluttered.

*   **Large:** Best for primary actions. Shows the label below the icon.
*   **Normal:** Best for secondary actions. Shows the label to the right of a small icon.

```xml
<group id="SizeDemo" label="Size Example">
    <button id="btnLrg" label="Primary" size="large" imageMso="FileSave" onAction="MyMacro" />
    <button id="btnSml" label="Secondary" size="normal" imageMso="FileSave" onAction="MyMacro" />
</group>
```

---

## 3. Native Icons (`imageMso`)
Microsoft provides thousands of built-in icons. You do not need to ship image files with your add-in if you use `imageMso`.

**Popular `imageMso` names:**
*   **Navigation:** `StepForward`, `StepBackward`, `GoToNextRecord`
*   **Data:** `ChartTypeColumnClustered`, `Filter`, `DatabaseInsert`
*   **Status:** `SymbolCheck`, `UnknownStatus`, `ErrorChecking`
*   **Shapes:** `HappyFace`, `Sun`, `Heart`

*Tip: Search for "Office 365 imageMso gallery" online to find visual lists of all available icons.*

---

## 4. Advanced Control Types
Beyond the standard button, you can use specialized controls to save space or gather input.

### Separators
Add a vertical line between buttons to create logical clusters.
```xml
<button id="btn1" label="Add" imageMso="AddAccount" onAction="MyMacro" />
<separator id="sep1" />
<button id="btn2" label="Delete" imageMso="Delete" onAction="MyMacro" />
```

### Menus (Drop-downs)
Use menus to group related macros under a single button.
```xml
<menu id="MyMenu" label="Export Options" imageMso="ExportExcel" size="large">
    <button id="exportPdf" label="Export to PDF" onAction="ExportMacro" />
    <button id="exportCsv" label="Export to CSV" onAction="ExportMacro" />
</menu>
```

### Toggle Buttons
A button that stays "pressed" or "unpressed"—useful for turning settings on/off.
```xml
<toggleButton id="toggleGrid" label="Gridlines" getPressed="GetPressedMacro" onAction="ToggleAction" imageMso="ViewGridlines" />
```

---

## 5. ScreenTips and SuperTips
Improve your Add-in's usability by adding documentation that appears when a user hovers over a button.

*   **Screentip:** The bold title of the tooltip.
*   **Supertip:** The detailed description paragraph.

```xml
<button id="btnInfo" 
        label="Process Data" 
        onAction="MyMacro" 
        imageMso="CalculateNow" 
        screentip="Runs the Data Engine" 
        supertip="This macro will validate all rows in the active sheet, remove duplicates, and generate a summary report." />
```

---

## Summary Checklist for Chapter 2
1.  **Unique IDs:** Every element (`tab`, `group`, `button`) must have a unique `id`.
2.  **Case Sensitivity:** XML tags and attributes (like `onAction`) are case-sensitive.
3.  **Layout Logic:** Use `size="large"` for your 1-2 most important features per group.
