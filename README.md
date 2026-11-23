
Option Explicit

'=========================================================
' CONFIGURATION
'=========================================================

Private Const COMPONENT_FOLDER As String = _
    "D:\VBA\VBA Components\"

Private Const ERR_BASE As Long = vbObjectError + 2000

Private Function GetComponentFolder() As String

    GetComponentFolder = COMPONENT_FOLDER

End Function

'=========================================================
' COMPONENT CONFIGURATION
'=========================================================

Private Function GetComponentNames() As Variant

    GetComponentNames = Array( _
        "modChangeCapture", _
        "modChangeLogger", _
        "modChangeUndo", _
        "clsChangeUndoRecord")

End Function

Private Function GetComponentFiles() As Variant

    GetComponentFiles = Array( _
        "modChangeCapture.bas", _
        "modChangeLogger.bas", _
        "modChangeUndo.bas", _
        "clsChangeUndoRecord.cls")

End Function

'=========================================================
' FIND EXISTING COMPONENT
'=========================================================

Private Function FindComponent( _
    ByVal targetWB As Workbook, _
    ByVal componentName As String) As Object

    Dim component As Object

    For Each component In targetWB.VBProject.VBComponents

        If StrComp(component.Name, componentName, _
                   vbTextCompare) = 0 Then

            Set FindComponent = component
            Exit Function

        End If

    Next component

End Function

'=========================================================
' VALIDATE COMPONENT FILES
'=========================================================

Public Function ValidateLoggerComponents() As Boolean

    Dim folderPath As String
    Dim componentFiles As Variant
    Dim fileName As Variant
    Dim missingFiles As String
    Dim i As Long

    On Error GoTo ErrorHandler

    folderPath = GetComponentFolder()

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    componentFiles = GetComponentFiles()

    For i = LBound(componentFiles) To UBound(componentFiles)

        fileName = folderPath & componentFiles(i)

        If Len(Dir$(CStr(fileName))) = 0 Then

            missingFiles = missingFiles & vbCrLf & _
                           " - " & componentFiles(i)

        End If

    Next i

    If Len(missingFiles) > 0 Then

        MsgBox "Missing component files:" & vbCrLf & _
               missingFiles & vbCrLf & vbCrLf & _
               "Expected folder:" & vbCrLf & folderPath, _
               vbExclamation, "Component Validation"

        Exit Function

    End If

    ValidateLoggerComponents = True

    Exit Function

ErrorHandler:

    MsgBox "Component validation error:" & vbCrLf & _
           Err.Description, _
           vbCritical, "Component Validation"

End Function

'=========================================================
' DEPLOY LOGGER COMPONENTS
'=========================================================

Public Function DeployVBAComponents( _
    ByVal targetWB As Workbook) As Boolean

    Dim componentFolder As String
    Dim backupFolder As String

    Dim componentNames As Variant
    Dim componentFiles As Variant

    Dim backupFiles As Object
    Dim component As Object
    Dim importedComponent As Object

    Dim i As Long
    Dim componentName As String
    Dim componentPath As String
    Dim backupPath As String

    Dim importStarted As Boolean
    Dim rollbackErrors As String
    Dim originalError As String

    On Error GoTo ErrorHandler

    '-----------------------------------------------------
    ' 1. Validate destination
    '-----------------------------------------------------

    If targetWB Is Nothing Then
        Err.Raise ERR_BASE + 1, , _
            "Destination workbook was not supplied."
    End If

    If targetWB.ReadOnly Then
        Err.Raise ERR_BASE + 2, , _
            "Destination workbook is read-only."
    End If

    If targetWB.IsAddin Then
        Err.Raise ERR_BASE + 3, , _
            "Add-in workbooks are not supported."
    End If

    '-----------------------------------------------------
    ' 2. Validate source files
    '-----------------------------------------------------

    componentFolder = GetComponentFolder()

    If Right$(componentFolder, 1) <> "\" Then
        componentFolder = componentFolder & "\"
    End If

    componentNames = GetComponentNames()
    componentFiles = GetComponentFiles()

    If Not ValidateLoggerComponents() Then
        Err.Raise ERR_BASE + 4, , _
            "Logger component validation failed."
    End If

    '-----------------------------------------------------
    ' 3. Create backup folder
    '-----------------------------------------------------

    backupFolder = Environ$("TEMP") & _
                   "\VBA_Logger_Backup_" & _
                   Format$(Now, "yyyymmdd_hhnnss") & _
                   "_" & Format$(CLng(Timer * 1000), "00000000") & "\"

    MkDir Left$(backupFolder, Len(backupFolder) - 1)

    Set backupFiles = CreateObject("Scripting.Dictionary")

    '-----------------------------------------------------
    ' 4. Export existing logger components
    '-----------------------------------------------------

    For i = LBound(componentNames) To UBound(componentNames)

        componentName = CStr(componentNames(i))

        Set component = FindComponent(targetWB, componentName)

        If Not component Is Nothing Then

            If i = 3 Then
                backupPath = backupFolder & componentName & ".cls"
            Else
                backupPath = backupFolder & componentName & ".bas"
            End If

            component.Export backupPath

            backupFiles.Add componentName, backupPath

        End If

        Set component = Nothing

    Next i

    '-----------------------------------------------------
    ' 5. Remove only matching logger components
    '-----------------------------------------------------

    For i = LBound(componentNames) To UBound(componentNames)

        componentName = CStr(componentNames(i))

        Set component = FindComponent(targetWB, componentName)

        If Not component Is Nothing Then
            targetWB.VBProject.VBComponents.Remove component
        End If

        Set component = Nothing

    Next i

    importStarted = True

    '-----------------------------------------------------
    ' 6. Import updated logger components
    '-----------------------------------------------------

    For i = LBound(componentFiles) To UBound(componentFiles)

        componentPath = componentFolder & _
                        CStr(componentFiles(i))

        Set importedComponent = _
            targetWB.VBProject.VBComponents.Import(componentPath)

        If importedComponent Is Nothing Then

            Err.Raise ERR_BASE + 5, , _
                "Import failed: " & componentFiles(i)

        End If

        If StrComp(importedComponent.Name, _
                   CStr(componentNames(i)), _
                   vbTextCompare) <> 0 Then

            Err.Raise ERR_BASE + 6, , _
                "Unexpected imported component name: " & _
                importedComponent.Name

        End If

        Set importedComponent = Nothing

    Next i

    DeployVBAComponents = True

    Exit Function

'=========================================================
' ERROR HANDLING AND ROLLBACK
'=========================================================

ErrorHandler:

    originalError = Err.Description

    'Remove partially imported components.

    If importStarted Then

        On Error Resume Next

        For i = LBound(componentNames) To UBound(componentNames)

            Set component = FindComponent( _
                targetWB, CStr(componentNames(i)))

            If Not component Is Nothing Then

                targetWB.VBProject.VBComponents.Remove component

            End If

            Set component = Nothing

        Next i

        On Error GoTo 0

    End If

    'Restore previously exported logger components.

    If Not backupFiles Is Nothing Then

        For i = LBound(componentNames) To UBound(componentNames)

            componentName = CStr(componentNames(i))

            If backupFiles.Exists(componentName) Then

                On Error Resume Next

                Set importedComponent = _
                    targetWB.VBProject.VBComponents.Import( _
                        CStr(backupFiles(componentName)))

                If Err.Number <> 0 Then

                    rollbackErrors = rollbackErrors & _
                        vbCrLf & componentName & _
                        ": " & Err.Description

                    Err.Clear

                End If

                Set importedComponent = Nothing

                On Error GoTo 0

            End If

        Next i

    End If

    If Len(rollbackErrors) > 0 Then

        originalError = originalError & vbCrLf & vbCrLf & _
            "WARNING: Rollback was incomplete:" & _
            rollbackErrors

    End If

    DeployVBAComponents = False

    MsgBox "VBA component deployment failed." & _
           vbCrLf & vbCrLf & originalError, _
           vbCritical, "Deployment Error"

End Function
