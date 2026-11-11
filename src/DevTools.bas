Attribute VB_Name = "DevTools"
Option Explicit

' Sync VBA modules between this workbook and the src\ folder next to it.
' Requires: File > Options > Trust Center > Trust Center Settings >
'           Macro Settings > "Trust access to the VBA project object model".
' DevTools itself is never overwritten by ImportSrc (re-import it manually if it changes).

Private Const SRC_DIR As String = "src"
Private Const SELF_NAME As String = "DevTools"

Public Sub ImportSrc()
    Dim proj As Object, comp As Object, srcPath As String, f As String
    Dim names As New Collection, nm As Variant

    srcPath = ThisWorkbook.Path & "\" & SRC_DIR & "\"
    Set proj = ThisWorkbook.VBProject

    f = Dir(srcPath & "*.bas")
    Do While Len(f) > 0
        names.Add Left$(f, Len(f) - 4)
        f = Dir()
    Loop

    For Each nm In names
        If StrComp(nm, SELF_NAME, vbTextCompare) <> 0 Then
            Set comp = Nothing
            On Error Resume Next
            Set comp = proj.VBComponents(CStr(nm))
            On Error GoTo 0
            If Not comp Is Nothing Then
                comp.Name = comp.Name & "_old"   ' rename first: Remove is deferred until the macro ends
                proj.VBComponents.Remove comp
            End If
            proj.VBComponents.Import srcPath & nm & ".bas"
        End If
    Next nm
    ' remove standard modules that no longer exist in src\ (renamed / deleted)
    Dim stale As New Collection, keep As Boolean, removed As Long
    For Each comp In proj.VBComponents
        If comp.Type = 1 And StrComp(comp.Name, SELF_NAME, vbTextCompare) <> 0 Then
            keep = False
            For Each nm In names
                If StrComp(nm, comp.Name, vbTextCompare) = 0 Then keep = True
            Next nm
            If Not keep Then stale.Add comp
        End If
    Next comp
    For Each comp In stale
        proj.VBComponents.Remove comp
        removed = removed + 1
    Next comp
    MsgBox names.Count & " file(s) imported from " & srcPath & IIf(removed > 0, vbCrLf & removed & " old module(s) removed", ""), _
           vbInformation, "DevTools"
End Sub

Public Sub ExportSrc()
    Dim comp As Object, srcPath As String, n As Long
    srcPath = ThisWorkbook.Path & "\" & SRC_DIR & "\"
    If Len(Dir(srcPath, vbDirectory)) = 0 Then MkDir srcPath
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then            ' vbext_ct_StdModule
            comp.Export srcPath & comp.Name & ".bas"
            n = n + 1
        End If
    Next comp
    MsgBox n & " module(s) exported to " & srcPath, vbInformation, "DevTools"
End Sub
