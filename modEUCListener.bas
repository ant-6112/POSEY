Attribute VB_Name = "modEUCListener"
Option Explicit

' ── State ─────────────────────────────────────────────────────────────────
Private m_Recording   As Boolean
Private m_MacroName   As String
Private m_ModuleName  As String

' ── Public entry points ───────────────────────────────────────────────────

Public Sub StartListener()
    If m_Recording Then
        MsgBox "Listener is already running.", vbInformation, "EUC Listener"
        Exit Sub
    End If

    ' Ask user for a name for the macro to record
    Dim sName As String
    sName = InputBox("Enter a name for the recorded macro:" & vbCrLf & _
                     "(letters/numbers/underscores only, no spaces)", _
                     "EUC Listener — Start", "RecordedMacro")
    If sName = "" Then Exit Sub

    ' Sanitise name
    sName = SanitiseName(sName)
    m_MacroName  = sName
    m_ModuleName = "Recorded_" & sName

    ' Visual feedback on the sheet
    With ThisWorkbook.Sheets("Listener")
        .Range("B5").Value  = "Status:  RECORDING — " & m_MacroName
        .Range("B5").Font.Color = RGB(192, 0, 0)
        .Range("B7").Value  = "Started: " & Now()
        .Range("B8").Value  = ""
    End With

    ' Start Excel macro recorder
    Application.MacroRecorder.RecordMacro m_MacroName, ""
    m_Recording = True
End Sub

Public Sub StopListener()
    If Not m_Recording Then
        MsgBox "Listener is not currently recording.", vbInformation, "EUC Listener"
        Exit Sub
    End If

    ' Stop the recorder
    Application.MacroRecorder.StopRecording
    m_Recording = False

    ' Give Excel a moment to write the module
    Application.Wait Now + TimeValue("00:00:01")

    ' Export the generated module to the workbook's folder
    Dim sExportPath As String
    sExportPath = ExportRecordedModule()

    ' Update UI
    With ThisWorkbook.Sheets("Listener")
        .Range("B5").Value  = "Status:  IDLE"
        .Range("B5").Font.Color = RGB(0, 128, 0)
        .Range("B8").Value  = "Stopped: " & Now()
        If sExportPath <> "" Then
            .Range("B10").Value = "Exported: " & sExportPath
        End If
    End With

    If sExportPath <> "" Then
        MsgBox "Recording saved!" & vbCrLf & vbCrLf & _
               "Module : " & m_ModuleName & vbCrLf & _
               "File   : " & sExportPath, _
               vbInformation, "EUC Listener — Done"
    Else
        MsgBox "Recording stopped." & vbCrLf & _
               "No module named '" & m_ModuleName & "' was found to export." & vbCrLf & _
               "Check the VBA editor (Alt+F11) for the recorded code.", _
               vbExclamation, "EUC Listener"
    End If
End Sub

' ── Helpers ───────────────────────────────────────────────────────────────

Private Function ExportRecordedModule() As String
    Dim vbProj  As Object   ' VBIDE.VBProject
    Dim vbComp  As Object   ' VBIDE.VBComponent
    Dim sFolder As String
    Dim sFile   As String

    On Error GoTo ErrHandler

    ' Workbook folder (fall back to Desktop if not yet saved)
    If ThisWorkbook.Path <> "" Then
        sFolder = ThisWorkbook.Path & Application.PathSeparator
    Else
        sFolder = Environ("USERPROFILE") & "\Desktop"
    End If

    sFile = sFolder & m_ModuleName & ".bas"

    Set vbProj = ThisWorkbook.VBProject

    ' Search for the module Excel just created
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = m_ModuleName Or _
           InStr(1, vbComp.Name, "Module", vbTextCompare) > 0 Then
            ' Try to rename if needed
            If vbComp.Name <> m_ModuleName Then
                On Error Resume Next
                vbComp.Name = m_ModuleName
                On Error GoTo ErrHandler
            End If
            vbComp.Export sFile
            ExportRecordedModule = sFile
            Exit Function
        End If
    Next vbComp

    ' Fallback: export the last standard module added
    Dim lastMod As Object
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 1 Then   ' vbext_ct_StdModule = 1
            Set lastMod = vbComp
        End If
    Next vbComp

    If Not lastMod Is Nothing Then
        lastMod.Name = m_ModuleName
        lastMod.Export sFile
        ExportRecordedModule = sFile
    End If

    Exit Function
ErrHandler:
    ExportRecordedModule = ""
End Function

Private Function SanitiseName(s As String) As String
    Dim i As Integer, c As String, result As String
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "[A-Za-z0-9_]" Then
            result = result & c
        End If
    Next i
    If result = "" Or Not (Left(result, 1) Like "[A-Za-z_]") Then
        result = "Macro_" & result
    End If
    SanitiseName = result
End Function
