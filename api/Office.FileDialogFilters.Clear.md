---
title: FileDialogFilters.Clear method (Office)
keywords: vbaof11.chm255005
f1_keywords:
- vbaof11.chm255005
ms.prod: office
api_name:
- Office.FileDialogFilters.Clear
ms.assetid: 1d5fa55e-6a61-d808-51a4-86116420f89f
ms.date: 01/09/2019
localization_priority: Normal
---


# FileDialogFilters.Clear method (Office)

Removes all filters currently applied in a file dialog box.

> [!NOTE] 
> The **Clear** method only works for the File Picker and File Open dialogs. This methods does **not** work when applied to the **Save As** and  **Folder Picker** objects. For example, **Application.FileDialog([msoFileDialogSaveAs](office.msofiledialogtype.md)).Filters.Clear** will result in a run-time error.
>

## Syntax

_expression_.**Clear**

_expression_ A variable that represents a **[FileDialogFilters](Office.FileDialogFilters.md)** object.


## Example

The following example clears the filters followed by adding one filter:

```vb
Sub sTestClear()
'----------------------------------------------------------------------------------------
' The msoFileDialogSaveAs dialog does NOT support file filters
'----------------------------------------------------------------------------------------

    Dim fd As FileDialog
    Dim tWbkFullName As String

    tWbkFullName = Application.ThisWorkbook.FullName
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    'Set fd = Application.FileDialog(msoFileDialogOpen)
    'Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .InitialFileName = tWbkFullName
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "jpg files", "*.jpg"
        
        If .Show = -1 Then ' this pops up the MsoFileDialogType
            ' User pressed the Action key, do your code
        Else
            ' User cancelled save.
        End If
    End With
End Sub
```

The following example shows how to set a prefered file type in the SaveAs dialog:

```vb
Sub sSetPreferedSaveExtension()

    Dim tWbkFullName As String
    Dim tI As Long
    Dim fd As FileDialog
    Dim fdFilterObj As FileDialogFilter

    tWbkFullName = Application.ThisWorkbook.FullName
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .InitialFileName = tWbkFullName
        .AllowMultiSelect = False
        ' --- .Clear, .Add & .Delete do not work with msoFileDialogSaveAs,
        '     Therefore, scroll your prefered SaveAs file extension into focus
        tI = 0
        For Each fdFilterObj In .Filters
            tI = tI + 1
            If fdFilterObj.Extensions = "*.xlsb" Then
                .FilterIndex = tI
                Exit For  ' dropdown now at the right SaveAsType
            End If
        Next fdFilterObj
        
        If .Show = -1 Then ' this pops up the FileSaveAs Dialogue
            ' User pressed the Action key, do your code
        Else
            ' User cancelled save.
        End If
    End With
End Sub
```


## See also

- [FileDialogFilters object members](overview/library-reference/filedialogfilters-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
