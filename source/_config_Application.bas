Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _config_Application
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const m_APPLICATIONVERSION As String = "0.4.0"

Private Const m_ApplicationName As String = "ACLib StructReader"
Private Const m_ApplicationFullName As String = m_ApplicationName
Private Const m_ApplicationTitle As String = m_ApplicationName
Private Const m_ApplicationIconFile As String = m_ApplicationName & ".ico"

Private Const m_DefaultErrorHandlerMode = ACLibErrorHandlerMode.aclibErrMsgBox

Private Const m_ApplicationStartFormName As String = "ACLibCodeModuleStruct"

'Public Const EXPORTFILEEXTENSION As String = ".html"

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">Möglichkeit einer Referenzübergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional oCurrentAppHandler As ApplicationHandler = Nothing)

'----------------------------------------------------------------------------
' Fehlerbehandlung
'
   modErrorHandler.DefaultErrorHandlerMode = m_DefaultErrorHandlerMode

'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   
   
'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If oCurrentAppHandler Is Nothing Then
      Set oCurrentAppHandler = CurrentApplication
   End If

   With oCurrentAppHandler
   
      'Anwendungsname
      .ApplicationName = m_ApplicationName
      .ApplicationFullName = m_ApplicationFullName
      
      'Titelleiste der Anwendung
      .ApplicationTitle = m_ApplicationTitle
      
      .Version = m_APPLICATIONVERSION
      
      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = m_ApplicationStartFormName

   End With
   
'----------------------------------------------------------------------------
' Erweiterung: ...
'
   'Konfiguration/Add-In-Einstellungen
   'modApplication.AddApplicationHandlerExtension New ACLibConfiguration
   
   'Import/Export von Dateien bzw. Access-Objekten
   'modApplication.AddApplicationHandlerExtension New ACLibFileManager
   
   'AppFile
   'modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile

'----------------------------------------------------------------------------
' Konfiguration nach Initialisierung der Erweiterungen
'
   'Icon der Anwendung und Fenster - erst nach AppFile-Initialisierung laden,
   '                                 falls Icon in AppFile-Tabelle enthalten ist.
   'oCurrentAppHandler.SetAppIcon CurrentProject.Path & "\" & m_ApplicationIconFile, True
   
   
End Sub

'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
Private Sub setAppFiles()
   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & m_ApplicationIconFile)
   'Call CurrentApplication.GetExtension("AppFile").SaveAppFile("GRAPHVIZ", CodeProject.Path & "\graphviz.zip")
End Sub
