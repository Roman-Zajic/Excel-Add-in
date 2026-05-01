# Manual Excel Ribbon Add-In Boilerplate

This repository demonstrates how to create a modern Excel Add-in (`.xlam`) with a custom Ribbon UI by manually editing the file's XML structure.

## 1. Create the Add-In & Macro
1. Open Excel and create a new workbook.
2. Press `Alt + F11` to open the VBA Editor.
3. Insert a new Module (`Insert > Module`) and paste the following:
```vba
' The (control As IRibbonControl) parameter is mandatory for Ribbon callbacks
Public Sub MyRibbonMacro(control As IRibbonControl)
    MsgBox "Add-in macro executed successfully!", vbInformation, "Success"
End Sub
```
4. Save the file as **MyAddIn.xlam**.
5. Close Excel.

## 2. Create the Custom UI
1. Rename `MyAddIn.xlam` to `MyAddIn.zip`.
2. Open the ZIP and create a new folder named `customUI`.
3. Inside the `customUI` folder, create a file named `customUI14.xml` with this content:
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="CustomTab" label="MY TOOLS">
        <group id="Group1" label="General">
          <button id="Btn1" label="Run Macro" size="large" onAction="MyRibbonMacro" imageMso="HappyFace" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

## 3. Link the UI in Relationships
1. Inside the ZIP, open `_rels/.rels`.
2. Add this line inside the `<Relationships>` tag before the closing `</Relationships>`:
```xml
<Relationship Id="R12345" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
```

## 4. Define Content Types
1. In the root of the ZIP, open `[Content_Types].xml`.
2. Add this line inside the `<Types>` tag:
```xml
<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
```

## 5. Finalize
1. Save and update all files inside the ZIP.
2. Rename `MyAddIn.zip` back to `MyAddIn.xlam`.
3. In Excel, go to **File > Options > Add-ins > Go...** and browse for your file.
4. The **"MY TOOLS"** tab will now appear in your Ribbon.

## Troubleshooting
If the tab does not appear, enable error reporting:
- **File > Options > Advanced > General > Show add-in user interface errors**.
