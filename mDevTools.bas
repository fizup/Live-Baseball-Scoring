Attribute VB_Name = "mDevTools"
Option Explicit
' -----------------------------------------------------------------------
' WICHTIG: Damit der Code auf das VBA-Projekt
' zugreifen kann, muss der Zugriff erlaubt sein:
' Geh zu Datei > Optionen > Trust Center (Sicherheitscenter)
' Einstellungen für das Trust Center:
' Makroeinstellungen -> Zugriff auf das VBA-Projektobjektmodell vertrauen
' -----------------------------------------------------------------------

' -----------------------------------------------------------------------
' Dieser Code löscht alte Module rigoros und
' erzwingt das Aufräumen des Speichers (DoEvents),
' bevor der Import startet. Das verhindert typische Namenskonflikte
' beim schnellen Hin- und Herwechseln zwischen VS Code und Excel

' Tastenkürzel für FastImport: Strg + Shift + I
' -----------------------------------------------------------------------
Sub FastImportFromVSCode()
Attribute FastImportFromVSCode.VB_ProcData.VB_Invoke_Func = "I\n14"
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim exportPath As String
    Dim compName As String
    Dim ext As String
    Dim vbComp As Object
    
    exportPath = ThisWorkbook.Path & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(exportPath)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. SCHRITT: Altes Zeug sauber löschen (außer diesem Entwickler-Modul)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then
            ' Verhindert, dass sich das Import-Skript selbst löscht
            If vbComp.Name <> "mDevTools" Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
        End If
    Next vbComp
    
    ' Dem Excel-Speicher kurz Zeit geben, die Löschung zu verarbeiten
    DoEvents
    
    ' 2. SCHRITT: Frischen Code aus dem VS Code Ordner laden
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        compName = fso.GetBaseName(file.Name)
        
        ' Nur importieren, wenn es nicht das laufende Entwickler-Modul selbst ist
        If compName <> "mDevTools" Then
            If ext = "bas" Or ext = "cls" Or ext = "frm" Then
                ThisWorkbook.VBProject.VBComponents.Import file.Path
            End If
        End If
    Next file
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Kurze unaufdringliche Statusmeldung in der Excel-Leiste unten links
    Application.StatusBar = "VBA-Code erfolgreich aus VS Code aktualisiert! (" & Time & ")"
End Sub

' -----------------------------------------------------------------------
' Exportiert alle VBA Komponenten in den
' Ordner in dem die Datei selbst liegt
' -----------------------------------------------------------------------
Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim fileExt As String
    
    ' Pfad des aktuellen Arbeitsmappen-Ordners abrufen
    exportPath = ThisWorkbook.Path & "\"
    
    ' Prüfen, ob die Datei bereits gespeichert wurde
    If ThisWorkbook.Path = "" Then
        MsgBox "Bitte speichern Sie die Excel-Datei zuerst!", vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' Schleife durch alle Komponenten des VBA-Projekts
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        ' Bestimmen der Dateiendung basierend auf dem Komponententyp
        Select Case vbComp.Type
            Case 1 ' Standard Modul
                fileExt = ".bas"
            Case 2 ' Klassenmodul
                fileExt = ".cls"
            Case 3 ' UserForm
                fileExt = ".frm"
            Case Else
                ' Microsoft Excel Objekte (DieseArbeitsmappe, Tabellenblätter) überspringen
                fileExt = ""
        End Select
        
        ' Exportieren, wenn es sich um ein Modul, eine Klasse oder ein Formular handelt
        If fileExt <> "" Then
            fileName = exportPath & vbComp.Name & fileExt
            vbComp.Export fileName
        End If
    Next vbComp
    
    MsgBox "Alle VBA-Komponenten wurden erfolgreich exportiert!", vbInformation, "Erfolg"
End Sub

