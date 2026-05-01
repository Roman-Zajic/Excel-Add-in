
# Mastering the Excel Ribbon

Welcome to the complete guide on building modern Microsoft Excel Add-ins (`.xlam`) from scratch. 

## Project Overview
The goal of this project is to demystify how the Office Fluent UI works. By the end of this guide, you will be able to create custom tabs, groups, and interactive controls that call VBA macros to automate your daily workflow.

### Key Features
- **Pure OpenXML:** Learn the internal architecture of `.xlam` files.
- **Custom UI:** Create professional-looking Ribbon tabs.
- **Dynamic Logic:** Make buttons that react to workbook changes.
- **Zero Dependencies:** Use only a text editor and a ZIP utility.

---

## Table of Contents

1.  **[Chapter 1: The Architecture](./01%20Architecture.md)**  
    Learn how an Excel file is structured internally and how to "wire" your XML into the ZIP archive.
2.  **[Chapter 2: Layout & Visuals](./02%20Layout%20and%20Visuals.md)**  
    Understand the Ribbon hierarchy (Tabs > Groups > Controls) and how to use built-in Office icons (`imageMso`).
3.  **[Chapter 3: The Callback System](./03%20Callback%20System.md)**  
    Connect your XML buttons to VBA macros using the mandatory callback signatures.
4.  **[Chapter 4: The Dynamic Ribbon](./04%20The%20Dynamic%20Ribbon%20.md)**  
    Master the `Invalidate` command to refresh the UI and create buttons that enable/disable based on logic.
5.  **[Chapter 5: Advanced Controls](./05%20Advanced%20Control%20Types.md)**  
    Implement Menus, Checkboxes, EditBoxes, and SplitButtons to create complex user interfaces.
6.  **[Chapter 6: Reference Appendix](./06%20Reference%20Appendix.md)**  
    A cheat sheet for common `imageMso` IDs, XML snippets, and VBA callback signatures.

---
