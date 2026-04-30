# cDialog - Windows API Dialog Object Class

## Overview

`cDialog` is a Windows API-based VB/VBA dialog object class that provides encapsulation for commonly used dialog functions such as file open/save dialogs and folder selection dialogs.

**Author**: (woeoio) 215879458@qq.com
**Version**: v1.0.2.0

## Main Features

- ✅ **Open File Dialog** (`ShowOpen`)
- ✅ **Save File Dialog** (`ShowSave`)
- ✅ **Folder Selection Dialog** (`ShowBrowseForFolder`)
- ✅ **File Extension Filtering**
- ✅ **Multiple File Selection Support**
- ✅ **Common Filter Presets** (Images, Documents, Code, etc.)
- ✅ **File Path Processing Utility Methods**
- ✅ **File Opening Functionality** (`OpenFile`, `SelectAndOpenFile`)

## Properties

### Basic Properties

| Property      | Type   | Description            | Default Value |
| ------------- | ------ | ---------------------- | ------------- |
| `DialogTitle` | String | Dialog title           | `""`          |
| `InitialDir`  | String | Initial directory      | `""`          |
| `DefaultExt`  | String | Default file extension | `""`          |
| `FileName`    | String | Default filename       | `""`          |
| `Filter`      | String | File filter            | `""`          |

### Behavior Properties

| Property          | Type    | Description                    | Default Value |
| ----------------- | ------- | ------------------------------ | ------------- |
| `MultiSelect`     | Boolean | Allow multiple file selection  | `False`       |
| `OverwritePrompt` | Boolean | Prompt when overwriting files  | `True`        |
| `PathMustExist`   | Boolean | Path must exist                | `True`        |
| `FileMustExist`   | Boolean | File must exist                | `True`        |
| `HideReadOnly`    | Boolean | Hide read-only checkbox        | `True`        |
| `ReadOnly`        | Boolean | Open in read-only mode         | `False`       |
| `NoChangeDir`     | Boolean | Don't change current directory | `False`       |
| `CreatePrompt`    | Boolean | Prompt when creating files     | `False`       |
| `NewDialogStyle`  | Boolean | Use new dialog style           | `True`        |

## Methods

### Core Methods

#### `ShowOpen()` As Variant

Displays an open file dialog. Returns a single file path (string) or an array of file paths (multi-select mode).

#### `ShowSave()` As String

Displays a save file dialog. Returns the file path selected by the user.

#### `ShowBrowseForFolder()` As String

Displays a folder selection dialog. Returns the folder path selected by the user.

#### `SelectFiles()` As Collection

Displays an open file dialog and returns a `Collection` object, providing a unified interface for both single and multiple file selection.

**Features:**

- Returns a `Collection` object for both single and multi-select modes
- Use `Collection.Count` to check if any files were selected
- Use `For Each` to iterate through all selected files
- Returns an empty collection (Count = 0) when no files are selected

### Filter Methods

#### `AddFilter(strDescription, strExtension)`

Adds a file filter.

- `strDescription`: Filter description, e.g., `"Text Files (*.txt)"`
- `strExtension`: Extension pattern, e.g., `"*.txt"`

#### `ClearFilter()`

Clears all filters.

#### `SetCommonFilters()`

Sets common file filters (text files, all files).

#### `SetImageFilters()`

Sets image file filters (BMP, JPEG, PNG, GIF).

#### `SetDocumentFilters()`

Sets document file filters (Word, Excel, PowerPoint, PDF, text).

#### `SetCodeFilters()`

Sets code file filters (Basic, C/C++, C#, Java, Python, JavaScript, HTML, XML).

### File Operation Methods

#### `OpenFile(strFilePath, [strOperation])` As Long

Opens a specified file using the system default program.

- `strFilePath`: File path
- `strOperation`: Operation type, default is `"open"`

#### `SelectAndOpenFile()` As Long

Displays a file selection dialog and opens the selected file(s).

#### `Reset()`

Resets all properties to default values.

### Path Processing Methods

#### `GetFilePath(strFullPath)` As String

Extracts the file path from a full path (without filename).

#### `GetFileName(strFullPath)` As String

Extracts the filename from a full path (without path).

#### `GetFileExtension(strFullPath)` As String

Extracts the file extension from a full path (including the dot).

#### `GetFileNameWithoutExt(strFullPath)` As String

Extracts the filename from a full path (without extension).

## Usage Examples

### Example 1: Basic Open File Dialog

```vb
Sub BasicOpenFileDialog()
    Dim dlg As New cDialog

    ' Set dialog title
    dlg.DialogTitle = "Select a file"

    ' Set initial directory
    dlg.InitialDir = "C:\Users\Public\Documents"

    ' Set default file extension
    dlg.DefaultExt = "txt"

    ' Show dialog
    Dim strFile As String
    strFile = dlg.ShowOpen()

    If strFile <> "" Then
        MsgBox "You selected: " & strFile
    Else
        MsgBox "You cancelled the selection"
    End If
End Sub
```

### Example 2: Open File with Filters

```vb
Sub OpenFileWithFilter()
    Dim dlg As New cDialog

    ' Add filters
    dlg.AddFilter "Text Files (*.txt)", "*.txt"
    dlg.AddFilter "Word Documents (*.doc;*.docx)", "*.doc;*.docx"
    dlg.AddFilter "All Files (*.*)", "*.*"

    ' Set dialog title
    dlg.DialogTitle = "Select a file to open"

    ' Show dialog
    Dim strFile As String
    strFile = dlg.ShowOpen()

    If strFile <> "" Then
        MsgBox "Selected file: " & strFile & vbCrLf & _
               "Filename: " & dlg.GetFileName(strFile) & vbCrLf & _
               "Extension: " & dlg.GetFileExtension(strFile)
    End If
End Sub
```

### Example 3: Multiple File Selection

```vb
Sub OpenMultipleFiles()
    Dim dlg As New cDialog

    ' Enable multi-select
    dlg.MultiSelect = True

    ' Add filters
    dlg.AddFilter "Image Files (*.bmp;*.jpg;*.png)", "*.bmp;*.jpg;*.png"
    dlg.AddFilter "All Files (*.*)", "*.*"

    ' Set dialog title
    dlg.DialogTitle = "Select multiple image files"

    ' Show dialog
    Dim varFiles As Variant
    varFiles = dlg.ShowOpen()

    If IsArray(varFiles) Then
        Dim i As Long
        Dim strMsg As String

        strMsg = "You selected " & UBound(varFiles) - LBound(varFiles) + 1 & " files:" & vbCrLf & vbCrLf
        For i = LBound(varFiles) To UBound(varFiles)
            strMsg = strMsg & (i + 1) & ". " & varFiles(i) & vbCrLf
        Next i

        MsgBox strMsg
    ElseIf varFiles <> "" Then
        MsgBox "You selected: " & varFiles
    Else
        MsgBox "You cancelled the selection"
    End If
End Sub
```

### Example 4: Using Preset Filters

```vb
Sub UsePresetFilters()
    Dim dlg As New cDialog

    ' Use preset image filters
    dlg.SetImageFilters
    dlg.DialogTitle = "Select an image file"
    Dim strImage As String
    strImage = dlg.ShowOpen()

    ' Reset and use document filters
    dlg.Reset
    dlg.SetDocumentFilters
    dlg.DialogTitle = "Select a document file"
    Dim strDoc As String
    strDoc = dlg.ShowOpen()

    ' Reset and use code filters
    dlg.Reset
    dlg.SetCodeFilters
    dlg.DialogTitle = "Select a code file"
    Dim strCode As String
    strCode = dlg.ShowOpen()
End Sub
```

### Example 5: Save File Dialog

```vb
Sub SaveFileDialog()
    Dim dlg As New cDialog

    ' Set dialog title
    dlg.DialogTitle = "Save File"

    ' Set initial directory and default filename
    dlg.InitialDir = "C:\Users\Public\Documents"
    dlg.FileName = "New Document"

    ' Set default extension
    dlg.DefaultExt = "txt"

    ' Add filters
    dlg.AddFilter "Text Files (*.txt)", "*.txt"
    dlg.AddFilter "All Files (*.*)", "*.*"

    ' Disable overwrite prompt (optional)
    dlg.OverwritePrompt = False

    ' Show dialog
    Dim strFile As String
    strFile = dlg.ShowSave()

    If strFile <> "" Then
        MsgBox "Save to: " & strFile
        ' Add code to save the file here
    End If
End Sub
```

### Example 6: Folder Selection Dialog

```vb
Sub BrowseForFolderDialog()
    Dim dlg As New cDialog

    ' Set dialog title
    dlg.DialogTitle = "Select a folder"

    ' Show dialog
    Dim strFolder As String
    strFolder = dlg.ShowBrowseForFolder()

    If strFolder <> "" Then
        MsgBox "You selected folder: " & strFolder
    End If
End Sub
```

### Example 7: Select and Open File

```vb
Sub SelectAndOpenFileExample()
    Dim dlg As New cDialog

    ' Set filters
    dlg.AddFilter "Text Files (*.txt)", "*.txt"
    dlg.AddFilter "All Files (*.*)", "*.*"

    ' Select and open file
    Dim lngResult As Long
    lngResult = dlg.SelectAndOpenFile()

    If lngResult > 0 Then
        MsgBox "File opened"
    Else
        MsgBox "You cancelled or failed to open the file"
    End If
End Sub
```

### Example 8: Open Specific File Directly

```vb
Sub OpenSpecificFile()
    Dim dlg As New cDialog
    Dim strFile As String

    strFile = "C:\Users\Public\Documents\example.txt"

    ' Check if file exists
    If Dir(strFile) <> "" Then
        Dim lngResult As Long
        lngResult = dlg.OpenFile(strFile)

        If lngResult > 32 Then
            MsgBox "File opened successfully"
        Else
            MsgBox "Failed to open file"
        End If
    Else
        MsgBox "File does not exist: " & strFile
    End If
End Sub
```

### Example 9: Path Processing Utilities

```vb
Sub PathProcessingExample()
    Dim dlg As New cDialog
    Dim strFullPath As String

    strFullPath = "C:\Users\Public\Documents\Report.xlsx"

    Debug.Print "Full path: " & strFullPath
    Debug.Print "Folder path: " & dlg.GetFilePath(strFullPath)
    Debug.Print "Filename: " & dlg.GetFileName(strFullPath)
    Debug.Print "Filename (no extension): " & dlg.GetFileNameWithoutExt(strFullPath)
    Debug.Print "Extension: " & dlg.GetFileExtension(strFullPath)

    ' Output:
    ' Full path: C:\Users\Public\Documents\Report.xlsx
    ' Folder path: C:\Users\Public\Documents
    ' Filename: Report.xlsx
    ' Filename (no extension): Report
    ' Extension: .xlsx
End Sub
```

### Example 10: Using SelectFiles for Unified File Selection

```vb
Sub UseSelectFilesCollection()
    Dim dlg As New cDialog
    Dim files As Collection
    Dim file As Variant
    Dim strMsg As String

    ' Configure dialog
    dlg.DialogTitle = "Select files"
    dlg.SetImageFilters
    dlg.MultiSelect = True  ' Can be True or False

    ' Get file collection
    Set files = dlg.SelectFiles()

    ' Check if any files were selected
    If files.Count > 0 Then
        strMsg = "You selected " & files.Count & " file(s):" & vbCrLf & vbCrLf
        For Each file In files
            strMsg = strMsg & "- " & file & vbCrLf
        Next file
        MsgBox strMsg
    Else
        MsgBox "You cancelled the selection"
    End If
End Sub
```

**Notes:**

- Whether `MultiSelect` is `True` or `False`, it always returns a `Collection` object
- Use `files.Count` to check if any files were selected
- Use `For Each` to iterate through files, no need to distinguish between single and multi-select
- Simpler and more unified code

### Example 11: Complete File Processing Workflow

```vb
Sub CompleteFileWorkflow()
    Dim dlg As New cDialog
    Dim strFile As String
    Dim strContent As String
    Dim fso As Object
    Dim ts As Object

    ' Step 1: Select file
    dlg.SetCommonFilters
    dlg.DialogTitle = "Select a text file to read"
    strFile = dlg.ShowOpen()

    If strFile = "" Then
        MsgBox "You cancelled the selection"
        Exit Sub
    End If

    ' Step 2: Read file content
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(strFile) Then
        Set ts = fso.OpenTextFile(strFile, 1) ' 1 = ForReading
        strContent = ts.ReadAll
        ts.Close

        MsgBox "File content: " & vbCrLf & vbCrLf & strContent
    End If

    ' Step 3: Modify content and prepare for save
    strContent = strContent & vbCrLf & vbCrLf & "--- Edited at " & Now() & " ---"

    ' Step 4: Select save location
    dlg.Reset
    dlg.DialogTitle = "Save modified file"
    dlg.FileName = dlg.GetFileNameWithoutExt(strFile) & "_modified"
    dlg.DefaultExt = "txt"

    Dim strSavePath As String
    strSavePath = dlg.ShowSave()

    If strSavePath <> "" Then
        Set ts = fso.CreateTextFile(strSavePath, True) ' True = overwrite
        ts.Write strContent
        ts.Close

        MsgBox "File saved to: " & strSavePath
    End If
End Sub
```

## Notes

1. **Buffer Size**: A 32KB buffer is allocated in multi-select mode, sufficient for handling selection of a large number of files.

2. **Return Value Types**:
   - `ShowOpen()` returns `Variant`, which can be a string (single select) or an array (multi-select)
   - Check return value type using `IsArray()` or `VarType()` before use

3. **File Existence**: When `FileMustExist = True` is set, users can only select existing files.

4. **Path Processing**: All path processing methods use backslash `\` as the path separator.

5. **Multiple File Selection**: When `MultiSelect` is enabled, the return value is an array where the first element is the path and subsequent elements are filenames.

## Technical Details

This class uses the following Windows APIs:

- `GetOpenFileNameW` / `GetSaveFileNameW` - File dialogs
- `SHBrowseForFolder` - Folder selection dialog
- `SHGetPathFromIDList` - Path conversion
- `ShellExecute` - File opening

Supported Windows versions: Windows XP and later

## Version History

- **v1.0.2.0** - Added `SelectFiles()` method that returns a `Collection` object, providing unified handling for both single and multiple file selection scenarios
- **v1.0** - Initial version, supports basic file open/save and folder selection features

## License

Please add appropriate license information based on project requirements.

---

**Need Help?** Contact author: 215879458@qq.com

