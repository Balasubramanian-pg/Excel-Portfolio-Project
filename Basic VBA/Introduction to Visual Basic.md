# **Excel Favorites Ribbon Add-In – Detailed Reference Guide**

## **Glossary of Terms**

| Term     | Meaning                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **COM**  | **Component Object Model** (COM) is a binary-interface standard for software components introduced by Microsoft in 1993. It enables inter-process communication and dynamic object creation across a wide range of programming languages. COM is the foundation for many Microsoft technologies such as OLE, OLE Automation, ActiveX, COM+, DCOM, Windows Shell, DirectX, UMDF, and Windows Runtime. It allows different applications or components to work together seamlessly, even if written in different languages. |
| **VBA**  | **Visual Basic for Applications** is an implementation of Microsoft’s event-driven programming language Visual Basic 6. It uses the Visual Basic Runtime Library and is typically embedded inside host applications like Excel, Word, or Access. VBA can automate tasks, manipulate data, and even control one application from another via OLE Automation. VBA can consume (but not create) ActiveX/COM DLLs. Newer versions include support for class modules for object-oriented code structures.                     |
| **VSTO** | **Visual Studio Tools for Office** is a set of tools (available as project templates and runtime) for creating .NET-based customizations in Microsoft Office applications. Introduced for Office 2003 and later, VSTO lets Office applications host the .NET CLR, enabling developers to build rich, managed-code extensions. It supports both document-level and application-level add-ins.                                                                                                                             |
| **XML**  | **Extensible Markup Language** is a markup language that defines rules for encoding documents in a way that is both human-readable and machine-readable. XML is designed for simplicity, generality, and usability over the internet. It’s text-based, supports Unicode for multiple languages, and is widely used for data exchange, configuration files, and structured storage in software systems, including Office file formats and web services.                                                                   |

---

## **Overview of the “Favorites” Ribbon**

The **Favorites Ribbon** is a custom Excel ribbon tab added immediately after the **Home** tab when Excel launches.
It groups frequently-used commands into logical categories for quick access, reducing navigation time and providing one-click access to common Excel and Windows tools.

---

## **Ribbon Groups & Commands**

### 1. **Worksheet (Group)**

Commands related to saving and editing worksheet content.

* **Save**

  * **Action:** Saves the current workbook.
  * **Shortcut:** Ctrl + S
  * **Use Case:** Prevents data loss by committing changes to disk.
* **Save As**

  * **Action:** Saves the workbook under a new file name or location.
  * **Shortcut:** F12
  * **Use Case:** Create backup versions or save in different formats (XLSX, CSV, PDF).

---

### 2. **Edit (Group)**

Quick access to basic editing tools.

* **Undo**

  * **Action:** Reverts the last action.
  * **Shortcut:** Ctrl + Z
* **Copy**

  * **Action:** Copies the selected data to the clipboard.
  * **Shortcut:** Ctrl + C
* **Cut**

  * **Action:** Removes the selected data and places it on the clipboard.
  * **Shortcut:** Ctrl + X
* **Paste**

  * **Action:** Inserts clipboard contents into the active cell or selection.
  * **Shortcut:** Ctrl + V
* **Spelling**

  * **Action:** Checks the active sheet for spelling errors.
  * **Shortcut:** F7

---

### 3. **Print (Group)**

Print preparation and execution tools.

* **Setup**

  * **Action:** Opens the Page Setup dialog box (Sheet tab active).
  * **Use Case:** Configure print range, print titles, and gridline display.
* **Preview**

  * **Action:** Displays Print Preview.
  * **Shortcut:** Ctrl + F2
* **Print**

  * **Action:** Sends the current sheet or selection to the printer.
  * **Shortcut:** Ctrl + P

---

### 4. **Program (Group)**

File management and Excel settings.

* **New** – Creates a blank workbook.
* **Open** – Opens an existing file. Shortcut: Ctrl + O
* **Close** – Closes the active workbook.
* **Properties** – Opens the file properties dialog.
* **Options** – Opens Excel Options dialog for application settings.
* **Exit** – Closes Excel entirely.

---

### 5. **Evaluate (Group)**

Tools for quick calculations and recalculations.

* **Windows Calculator**

  * **Action:** Opens Windows Calculator in Standard mode.
  * **Extra Modes:** Scientific, Programmer, and Statistical.
* **Calculate Now**

  * **Action:** Forces recalculation of all formulas regardless of calculation mode.

---

### 6. **Annotate (Group)**

Capture and document parts of your work.

* **Excel Camera**

  * **Action:** Takes a live snapshot of selected cells, charts, or ranges.
  * **Special Feature:** The snapshot updates automatically when source data changes.
* **Snipping Tool**

  * **Action:** Launches Windows Snipping Tool for screen captures.
  * **Modes:** Free-form, Rectangular, Window, Full-screen.
* **Problem Step Recorder (PSR)**

  * **Action:** Records user actions step-by-step for troubleshooting.
  * **Output:** Saves as a .zip containing an .mht file with screenshots and text.

---

### 7. **Options (Group)**

**Add-In Settings**

* **VSTO Settings**

  * **Application Settings** – Fixed at development time; require redeployment to change.
  * **User Settings** – Can be modified by the user at runtime.

**VBA Settings**

* **Add a New Setting**

```vba
ThisWorkbook.CustomDocumentProperties.Add _
Name:="App_ReleaseDate", LinkToContent:=False, _
Type:=msoPropertyTypeDate, Value:="31-Jul-2017 1:05pm"
```

* **Update an Existing Setting**

```vba
ThisWorkbook.CustomDocumentProperties.Item("App_ReleaseDate").Value = "31-Jul-2017 1:05pm"
```

* **Delete a Setting**

```vba
ThisWorkbook.CustomDocumentProperties.Item("App_ReleaseDate").Delete
```

---

### 8. **Help (Group)**

* **How To…** – Opens a browser-based guide.
* **Report Issue** – Opens issue submission page in browser.
* **New Version Available** – Displays only if a newer add-in version is detected. Downloads and installs latest version.

---

### 9. **About (Group)**

* **Add-in Name** – Displays add-in name and version number.
* **Release Date** – Displays release date of the current version.
* **Copyright** – Displays author/organization.

---
