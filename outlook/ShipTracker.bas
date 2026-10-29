' ============================================
' SHIP TRACKER PRO — Outlook VBA Email Logger
' ============================================
' v2 - Fixed: better JSON cleaning, error reporting, handles all email types
'
' SETUP:
' 1. Open Outlook → Press Alt+F11
' 2. Left panel → "ThisOutlookSession"
' 3. Paste ALL this code → Ctrl+S
' 4. Add to Quick Access Toolbar:
'    File → Options → Quick Access Toolbar → Macros → LogToShipTracker → Add
'
' ★★★ CHANGE THIS TO YOUR NAME ★★★

Const MY_NAME = "CHANGE_ME"
Const FIREBASE_URL = "https://firestore.googleapis.com/v1/projects/ship-checker/databases/(default)/documents/incomingEmails"
Const CURRENT_VERSION = "3.7"
Const VBA_CODE_URL = ""
' End of Day email button
Const FIREBASE_WEB_KEY = "AIzaSyA2HSftodTHDq3CJbeN7O4i0u4cADawJwI"   ' public web key (same as the website)
Const EOD_DOC_URL = "https://firestore.googleapis.com/v1/projects/ship-checker/databases/(default)/documents/eodLogs/"

' ★★★ CHANGE THIS TO YOUR TEAMS/SHAREPOINT CONFIG FILE URL ★★★
'Config is now loaded from Firebase - no URL needed

Dim g_entryId As String
Dim g_threadGroup As String
Dim g_startupScanDone As Boolean
Dim g_lastConfigResponse As String
Dim g_emailsLoggedToday As Long
Dim g_lastShipLogged As String

' Dynamic config — loaded from Teams on startup
Dim g_configShips As String
Dim g_configKeywords As String
Dim g_configTeam As String
Dim g_configLoaded As Boolean
Dim g_configVersion As String

Private Sub LoadConfig()
    On Error Resume Next
    
    If g_configLoaded Then Exit Sub
    
    ' Read config from Firebase Firestore
    Dim configUrl As String
    configUrl = "https://firestore.googleapis.com/v1/projects/ship-checker/databases/(default)/documents/settings/vbaConfig"
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", configUrl, False
    http.Send
    
    If http.Status = 200 Then
        Dim resp As String
        resp = http.responseText
        g_lastConfigResponse = resp
        
        ' Parse version
        g_configVersion = ExtractFirestoreField(resp, "version")
        
        ' Parse ships
        g_configShips = ExtractFirestoreField(resp, "ships")
        
        ' Parse keywords
        g_configKeywords = ExtractFirestoreField(resp, "keywords")
        
        ' Parse team
        g_configTeam = ExtractFirestoreField(resp, "team")
        
        g_configLoaded = True
        Debug.Print Now & " | Config loaded from Firebase v" & g_configVersion
    Else
        Debug.Print Now & " | Config load failed: " & http.Status
        g_configLoaded = False
    End If
    
    Set http = Nothing
End Sub

' Extract a string field value from Firestore JSON response
Private Function ExtractFirestoreField(json As String, fieldName As String) As String
    On Error Resume Next
    
    ' Step 1: Find the field name
    Dim fieldPos As Long
    fieldPos = InStr(1, json, """" & fieldName & """")
    If fieldPos = 0 Then
        ExtractFirestoreField = ""
        Exit Function
    End If
    
    ' Step 2: Find "stringValue" after the field name
    Dim svPos As Long
    svPos = InStr(fieldPos, json, """stringValue""")
    If svPos = 0 Then
        ExtractFirestoreField = ""
        Exit Function
    End If
    
    ' Make sure this stringValue belongs to THIS field (not a later one)
    ' Check that no other field name appears between fieldPos and svPos
    Dim nextFieldCheck As Long
    nextFieldCheck = svPos - fieldPos
    If nextFieldCheck > 200 Then
        ExtractFirestoreField = ""
        Exit Function
    End If
    
    ' Step 3: Find the value after "stringValue": "
    Dim colonPos As Long
    colonPos = InStr(svPos, json, ":")
    If colonPos = 0 Then ExtractFirestoreField = "": Exit Function
    
    ' Find opening quote of value
    Dim openQuote As Long
    openQuote = InStr(colonPos, json, """")
    If openQuote = 0 Then ExtractFirestoreField = "": Exit Function
    openQuote = openQuote + 1
    
    ' Find closing quote
    Dim closeQuote As Long
    closeQuote = InStr(openQuote, json, """")
    If closeQuote = 0 Then ExtractFirestoreField = "": Exit Function
    
    ExtractFirestoreField = Mid(json, openQuote, closeQuote - openQuote)
End Function


' ============================================
' TEAM ROSTER — Maps usernames/emails to real names
' Add your team members here
' ============================================
Private Function GetRealName(senderName As String, senderEmail As String) As String
    Dim n As String: n = LCase(Trim(senderName))
    Dim e As String: e = LCase(Trim(senderEmail))
    
    ' Check Teams config first
    If g_configLoaded And Len(g_configTeam) > 5 Then
        Dim teamLines() As String
        teamLines = Split(Replace(Replace(g_configTeam, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        Dim tl As Long
        For tl = 0 To UBound(teamLines)
            Dim parts2() As String
            If InStr(teamLines(tl), "=") > 0 Then
                parts2 = Split(teamLines(tl), "=")
                Dim matchKey As String: matchKey = LCase(Trim(parts2(0)))
                Dim matchName As String: matchName = Trim(parts2(1))
                If Len(matchKey) > 0 Then
                    If InStr(n, matchKey) > 0 Or InStr(e, matchKey) > 0 Then
                        GetRealName = matchName
                        Exit Function
                    End If
                End If
            End If
        Next tl
    End If
    
    ' Daniel Brown
    If InStr(n, "dbrown") > 0 Or InStr(n, "daniel") > 0 Or InStr(n, "d.brown") > 0 Or _
       InStr(n, "brown, daniel") > 0 Or InStr(n, "daniel brown") > 0 Or InStr(n, "daniel k") > 0 Or _
       InStr(e, "dbrown") > 0 Or InStr(e, "daniel.brown") > 0 Or InStr(e, "d.brown") > 0 Or _
       InStr(e, "daniel.k.brown") > 0 Then
        GetRealName = "Daniel Brown": Exit Function
    End If
    
    ' Clyde Grant
    If InStr(n, "clyde") > 0 Or InStr(n, "grant, clyde") > 0 Or InStr(n, "clyde grant") > 0 Or _
       InStr(n, "c.grant") > 0 Or InStr(n, "cgrant") > 0 Or _
       InStr(e, "clyde") > 0 Or InStr(e, "cgrant") > 0 Or InStr(e, "c.grant") > 0 Then
        GetRealName = "Clyde Grant": Exit Function
    End If
    
    ' Uriel Pena
    If InStr(n, "uriel") > 0 Or InStr(n, "up_reap") > 0 Or InStr(n, "u_pena") > 0 Or _
       InStr(n, "pena, uriel") > 0 Or InStr(n, "uriel pena") > 0 Or InStr(n, "u.pena") > 0 Or _
       InStr(e, "uriel") > 0 Or InStr(e, "up_reap") > 0 Or InStr(e, "u.pena") > 0 Or _
       InStr(e, "upena") > 0 Or InStr(e, "uriel.pena") > 0 Then
        GetRealName = "Uriel Pena": Exit Function
    End If
    
    ' Jayleel Kazimm
    If InStr(n, "jayleel") > 0 Or InStr(n, "jayreap") > 0 Or InStr(n, "jay_reap") > 0 Or _
       InStr(n, "kazimm") > 0 Or InStr(n, "j.kazimm") > 0 Or InStr(n, "jkazimm") > 0 Or _
       InStr(n, "hanif") > 0 Or InStr(n, "kaazim") > 0 Or InStr(n, "jaleel") > 0 Or _
       InStr(e, "jayleel") > 0 Or InStr(e, "jayreap") > 0 Or InStr(e, "jay_reap") > 0 Or _
       InStr(e, "kazimm") > 0 Or InStr(e, "j.kazimm") > 0 Or InStr(e, "hanif") > 0 Or _
       InStr(e, "kaazim") > 0 Then
        GetRealName = "Jayleel Kazimm": Exit Function
    End If
    
    ' Damian Pierman (Lead)
    If InStr(n, "pierman") > 0 Or InStr(n, "damian") > 0 Or _
       InStr(e, "pierman") > 0 Or InStr(e, "damian.j.pierman") > 0 Then
        GetRealName = "Damian Pierman": Exit Function
    End If
    
    ' Eryk Kedzierski
    If InStr(n, "kedzierski") > 0 Or InStr(n, "eryk") > 0 Or _
       InStr(e, "kedzierski") > 0 Or InStr(e, "eryk") > 0 Then
        GetRealName = "Eryk Kedzierski": Exit Function
    End If
    
    ' No match — return original name cleaned up
    ' Handle "Last, First" format
    If InStr(senderName, ",") > 0 Then
        Dim parts() As String
        parts = Split(senderName, ",")
        If UBound(parts) >= 1 Then
            GetRealName = Trim(parts(1)) & " " & Trim(parts(0))
            Exit Function
        End If
    End If
    
    GetRealName = senderName
End Function

' ============================================
' AUTO-KEYWORDS — Emails matching these get logged automatically
' Case-insensitive matching (checked as uppercase)
' ============================================
Private Function GetAutoKeywords() As String
    ' Use Teams config if loaded, otherwise built-in
    If g_configLoaded And Len(g_configKeywords) > 10 Then
        GetAutoKeywords = g_configShips & "," & g_configKeywords
        Exit Function
    End If
    
    ' Built-in fallback
    Dim k As String
    k = "DDG54,DDG55,DDG56,DDG57,DDG62,DDG63,DDG67,DDG70,DDG71,DDG72,DDG77,DDG82,DDG96,DDG98,DDG99,DDG101,DDG103,DDG109,DDG111,"
    k = k & "CG70,CG52,CG53,CG54,CG55,CG56,CG57,CG58,CG59,CG60,CG61,CG62,CG63,CG64,CG65,CG66,CG67,CG68,CG69,CG71,CG72,CG73,"
    k = k & "LCS1,LCS2,LCS3,LCS4,LCS5,LCS6,LCS7,LCS8,LCS9,LCS10,LCS11,LCS12,LCS13,LCS14,LCS15,LCS16,LCS17,LCS18,LCS19,LCS20,"
    k = k & "LPD17,LPD18,LPD19,LPD20,LPD21,LPD22,LPD23,LPD24,LPD25,LPD26,LPD27,LPD28,LPD29,LPD30,LPD31,LPD32,"
    k = k & "LHD1,LHD2,LHD3,LHD4,LHD5,LHD6,LHD7,LHD8,CVN68,CVN69,CVN70,CVN71,CVN72,CVN73,CVN74,CVN75,CVN76,CVN77,CVN78,CVN79,"
    k = k & "REAP,MINI REAP,FULL REAP,SATCOM,ROUTER,SWITCH,FIREWALL,NETWORK,TROUBLESHOOT,OUTAGE,OFFLINE,"
    k = k & "HARD DRIVE,HDD,SSD,POWER SUPPLY,BSOD,CRASH,SERVER,DIODE,TECHNICIAN,INSTALL,DEINSTALL"
    GetAutoKeywords = k
End Function


' ============================================
' STARTUP: Scan inbox for missed emails
' Fires when Outlook opens — catches overnight emails
' ============================================
Private Sub Application_Startup()
    ' Don't scan here — Outlook isn't connected to Exchange yet
    ' The scan will happen via MAPILogonComplete or first NewMailEx
    g_startupScanDone = False
End Sub

' Fires AFTER Outlook connects to Exchange — inbox is ready
Private Sub Application_MAPILogonComplete()
    On Error Resume Next
    
    ' Load config from Firebase
    LoadConfig
    
    ' Wait a few seconds for inbox to fully sync
    Dim waitUntil As Double
    waitUntil = Timer + 5
    Do While Timer < waitUntil
        DoEvents
    Loop
    
    g_startupScanDone = True
    RunCatchupScan
    
    ' Show persistent status bar
    Application.StatusBar = "Ship Tracker v" & CURRENT_VERSION & " Active | " & Format(Now, "m/d h:nn AM/PM")
End Sub

' Shared scan function — called by MAPILogonComplete and first NewMailEx
Private Sub RunCatchupScan()
    On Error Resume Next
    
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Dim inbox As Outlook.MAPIFolder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    
    ' Check emails from last 48 hours
    Dim cutoff As String
    cutoff = Format(DateAdd("h", -48, Now), "mm/dd/yyyy hh:nn AM/PM")
    
    Dim recentItems As Outlook.Items
    Set recentItems = inbox.Items.Restrict("[ReceivedTime] >= '" & cutoff & "'")
    
    If recentItems Is Nothing Then GoTo ScanSent
    If recentItems.Count = 0 Then GoTo ScanSent
    
    ' Sort by most recent first
    recentItems.Sort "[ReceivedTime]", True
    
    Dim scanned As Long: scanned = 0
    Dim logged As Long: logged = 0
    Dim i As Long
    
    For i = 1 To recentItems.Count
        If TypeOf recentItems.Item(i) Is Outlook.MailItem Then
            Dim m As Outlook.MailItem
            Set m = recentItems.Item(i)
            
            If MatchesKeywords(m) Then
                scanned = scanned + 1
                AutoLogEmail m
                logged = logged + 1
            End If
        End If
        
        If i > 200 Then Exit For
    Next i
    
    If logged > 0 Then
        Application.StatusBar = "Ship Tracker: Startup scan found " & logged & " emails | " & Format(Now, "h:nn AM/PM")
        Debug.Print Now & " | STARTUP SCAN: Scanned " & scanned & " | Logged " & logged
    End If

ScanSent:
    ' Also check Sent Items
    Dim sentFolder As Outlook.MAPIFolder
    Set sentFolder = ns.GetDefaultFolder(olFolderSentMail)
    Dim recentSent As Outlook.Items
    Set recentSent = sentFolder.Items.Restrict("[SentOn] >= '" & cutoff & "'")
    
    If Not recentSent Is Nothing Then
        recentSent.Sort "[SentOn]", True
        For i = 1 To recentSent.Count
            If TypeOf recentSent.Item(i) Is Outlook.MailItem Then
                Set m = recentSent.Item(i)
                If MatchesKeywords(m) Then
                    AutoLogEmail m
                End If
            End If
            If i > 200 Then Exit For
        Next i
    End If
End Sub

' ============================================
' AUTO-TRIGGER: Fires when NEW EMAIL arrives
' ============================================
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    On Error Resume Next
    
    ' First time any email arrives — run catchup scan for missed emails
    If Not g_startupScanDone Then
        g_startupScanDone = True
        RunCatchupScan
    End If
    
    ' Refresh ship/keyword config from Ship Tracker (lightweight - one small Firebase doc)
    g_configLoaded = False
    LoadConfig
    
    Dim ids() As String
    ids = Split(EntryIDCollection, ",")
    
    Dim i As Long
    For i = 0 To UBound(ids)
        Dim mail As Object
        Set mail = Application.Session.GetItemFromID(Trim(ids(i)))
        
        If Not mail Is Nothing Then
            If TypeOf mail Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = mail
                
                ' Check if email matches any keywords
                If MatchesKeywords(m) Then
                    AutoLogEmail m
                End If
            End If
        End If
    Next i
End Sub

' ============================================
' AUTO-TRIGGER: Fires when YOU SEND an email
' ============================================
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error Resume Next
    
    If TypeOf Item Is Outlook.MailItem Then
        Dim m As Outlook.MailItem
        Set m = Item
        
        If MatchesKeywords(m) Then
            ' Slight delay — send first, then log
            ' We log on a timer so the email sends without delay
            Dim logItem As Outlook.MailItem
            Set logItem = m.Copy
            AutoLogEmail logItem
        End If
    End If
End Sub

' ============================================
' CHECK KEYWORDS — does this email match?
' ============================================
Private Function MatchesKeywords(m As Outlook.MailItem) As Boolean
    Dim keywords() As String
    keywords = Split(GetAutoKeywords(), ",")
    
    Dim searchText As String
    searchText = UCase(SafeStr(m.subject) & " " & SafeStr(m.SenderName) & " " & SafeStr(m.SenderEmailAddress) & " " & SafeStr(Left(m.body, 500)))
    
    ' Also check first recipient
    If m.Recipients.Count > 0 Then
        searchText = searchText & " " & UCase(SafeStr(m.Recipients.Item(1).Name) & " " & SafeStr(m.Recipients.Item(1).Address))
    End If
    
    Dim k As Long
    For k = 0 To UBound(keywords)
        If Len(Trim(keywords(k))) > 0 Then
            If InStr(1, searchText, Trim(keywords(k))) > 0 Then
                MatchesKeywords = True
                Exit Function
            End If
        End If
    Next k
    
    MatchesKeywords = False
End Function

' ============================================
' AUTO-LOG — silently send email to Firebase
' ============================================
Private Sub AutoLogEmail(m As Outlook.MailItem)
    On Error Resume Next
    
    ' CHECK IF ALREADY LOGGED — prevents all duplicates
    Dim prop As Outlook.UserProperty
    Set prop = m.UserProperties.Find("ShipTrackerLogged")
    If Not prop Is Nothing Then
        Debug.Print Now & " | SKIP (already logged): " & Left(m.subject, 40)
        Exit Sub
    End If
    
    ' IGNORE LIST — skip automated/system emails
    Dim senderLower As String
    senderLower = LCase(m.SenderEmailAddress & " " & m.SenderName)
    If InStr(senderLower, "noreply") > 0 Or InStr(senderLower, "no-reply") > 0 Or _
       InStr(senderLower, "donotreply") > 0 Or InStr(senderLower, "do-not-reply") > 0 Or _
       InStr(senderLower, "mailer-daemon") > 0 Or InStr(senderLower, "postmaster") > 0 Or _
       InStr(senderLower, "newsletter") > 0 Or InStr(senderLower, "notification") > 0 Or _
       InStr(senderLower, "automated") > 0 Or InStr(senderLower, "system administrator") > 0 Or _
       InStr(senderLower, "helpdesk@") > 0 Or InStr(senderLower, "nmci") > 0 Or _
       InStr(senderLower, "global address") > 0 Or InStr(senderLower, "calendar") > 0 Or _
       InStr(senderLower, "sharepoint") > 0 Or InStr(senderLower, "teams") > 0 Then
        Debug.Print Now & " | SKIP (ignored sender): " & Left(m.SenderName, 30)
        Exit Sub
    End If
    
    ' SMART FILTER — skip internal team-to-team emails (no ship involved)
    Dim fromAddr As String: fromAddr = LCase(GetSMTPAddress(m, True))
    Dim isFromTeam As Boolean: isFromTeam = IsTeamMember(fromAddr)
    Dim allRecipientsTeam As Boolean: allRecipientsTeam = True
    Dim hasShipRecipient As Boolean: hasShipRecipient = False
    Dim ri As Long
    For ri = 1 To m.Recipients.Count
        Dim recipAddr As String
        recipAddr = LCase(GetRecipientSMTP(m.Recipients.Item(ri)))
        If Not IsTeamMember(recipAddr) Then allRecipientsTeam = False
        ' Check if any recipient is a ship email
        If InStr(recipAddr, "ddg") > 0 Or InStr(recipAddr, "@cg") > 0 Or _
           InStr(recipAddr, "@lcs") > 0 Or InStr(recipAddr, "@lpd") > 0 Or _
           InStr(recipAddr, "@lhd") > 0 Or InStr(recipAddr, "@cvn") > 0 Or _
           InStr(recipAddr, "sto@") > 0 Or InStr(recipAddr, "cso@") > 0 Or _
           InStr(recipAddr, "fco@") > 0 Then
            hasShipRecipient = True
        End If
    Next ri
    ' If FROM is team AND all recipients are team AND no ship in subject/body — skip
    If isFromTeam And allRecipientsTeam And Not hasShipRecipient Then
        Dim shipCheck As String
        shipCheck = DetectShip(SafeStr(m.subject) & " " & Left(SafeStr(m.body), 300))
        If shipCheck = "" Then
            Debug.Print Now & " | SKIP (internal team email): " & Left(m.subject, 40)
            Exit Sub
        End If
    End If
    
    Dim fromEmail As String: fromEmail = GetSMTPAddress(m, True)
    Dim fromName As String: fromName = GetRealName(SafeStr(m.SenderName), fromEmail)
    Dim toEmail As String: toEmail = ""
    Dim toName As String: toName = ""
    If m.Recipients.Count > 0 Then
        toEmail = GetRecipientSMTP(m.Recipients.Item(1))
        toName = GetRealName(SafeStr(m.Recipients.Item(1).Name), toEmail)
    End If
    
    Dim subject As String: subject = SafeStr(m.subject)
    Dim body As String: body = CleanEmailBody(m.body)
    Dim convId As String: convId = SafeStr(m.ConversationID)
    Dim recvAt As String
    On Error Resume Next
    If m.ReceivedTime > #1/1/2000# Then
        recvAt = Format(m.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z"
    ElseIf m.SentOn > #1/1/2000# Then
        recvAt = Format(m.SentOn, "yyyy-mm-ddThh:nn:ss") & "Z"
    Else
        recvAt = Format(Now, "yyyy-mm-ddThh:nn:ss") & "Z"
    End If
    On Error GoTo 0
    
    g_entryId = SafeStr(m.EntryID)
    
    Dim shipName As String
    shipName = DetectShip(subject & " " & Left(body, 500) & " " & fromName & " " & fromEmail & " " & toName & " " & toEmail)
    If shipName = "" Then shipName = "UNMATCHED"
    
    ' Silent send — no popups, no confirmation
    Dim result As String
    result = SendToFirebase(subject, body, fromName, fromEmail, toName, toEmail, convId, recvAt, shipName)
    
    ' Log to Immediate Window for debugging (Ctrl+G in VBA editor to see)
    ' Mark email as logged so it never gets sent again
    If result = "OK" Then
        On Error Resume Next
        Dim newProp As Outlook.UserProperty
        Set newProp = m.UserProperties.Add("ShipTrackerLogged", olText)
        newProp.Value = Format(Now, "yyyy-mm-ddThh:nn:ss")
        m.Save
        On Error GoTo 0
    End If
    
    ' Update status bar with running count
    If result = "OK" And shipName <> "UNMATCHED" Then
        g_emailsLoggedToday = g_emailsLoggedToday + 1
        g_lastShipLogged = shipName
        Application.StatusBar = "Ship Tracker v" & CURRENT_VERSION & " | Last: " & shipName & " - " & Left(subject, 30) & " | " & g_emailsLoggedToday & " logged today | " & Format(Now, "h:nn AM/PM")
    End If
    
    Debug.Print Now & " | AUTO-LOG: " & shipName & " | " & Left(subject, 50) & " | " & result
End Sub

' ============================================
' MAIN — Log single selected email
' ============================================
' ============================================
' ONE BUTTON — logs single email or full thread
' ============================================
' ============================================
' SCAN INBOX — run manually to catch missed emails
' ============================================
Public Sub ScanInbox()
    On Error Resume Next
    
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Dim inbox As Outlook.MAPIFolder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    
    Dim cutoff As String
    cutoff = Format(DateAdd("h", -48, Now), "mm/dd/yyyy hh:nn AM/PM")
    
    Dim recentItems As Outlook.Items
    Set recentItems = inbox.Items.Restrict("[ReceivedTime] >= '" & cutoff & "'")
    
    Dim logged As Long: logged = 0
    Dim checked As Long: checked = 0
    Dim i As Long
    
    If Not recentItems Is Nothing Then
        For i = 1 To recentItems.Count
            If TypeOf recentItems.Item(i) Is Outlook.MailItem Then
                Dim m As Outlook.MailItem
                Set m = recentItems.Item(i)
                If MatchesKeywords(m) Then
                    checked = checked + 1
                    AutoLogEmail m
                    logged = logged + 1
                End If
            End If
            If i > 200 Then Exit For
        Next i
    End If
    
    ' Also scan sent
    Dim sentFolder As Outlook.MAPIFolder
    Set sentFolder = ns.GetDefaultFolder(olFolderSentMail)
    Dim recentSent As Outlook.Items
    Set recentSent = sentFolder.Items.Restrict("[SentOn] >= '" & cutoff & "'")
    
    If Not recentSent Is Nothing Then
        For i = 1 To recentSent.Count
            If TypeOf recentSent.Item(i) Is Outlook.MailItem Then
                Set m = recentSent.Item(i)
                If MatchesKeywords(m) Then
                    checked = checked + 1
                    AutoLogEmail m
                    logged = logged + 1
                End If
            End If
            If i > 200 Then Exit For
        Next i
    End If
    
    MsgBox "Scan complete!" & vbCrLf & vbCrLf & _
           "Checked last 24 hours" & vbCrLf & _
           "Matched: " & checked & vbCrLf & _
           "Sent to Ship Tracker: " & logged & vbCrLf & vbCrLf & _
           "Duplicates will be filtered in the app.", vbInformation, "Ship Tracker Scan"
End Sub

Public Sub LogToShipTracker()
    Dim mail As Outlook.MailItem
    Dim rawBody As String
    Dim bodyLen As Long
    Dim fromPositions() As Long
    Dim sentPositions() As Long
    Dim pairPositions() As Long
    Dim fromCount As Long
    Dim sentCount As Long
    Dim pairCount As Long
    Dim pos As Long
    Dim fi As Long
    Dim si As Long
    Dim ci As Long
    Dim fp As Long
    Dim snippet As String
    Dim pairInfo As String
    Dim shipFE As String
    Dim shipTE As String
    Dim shipName As String
    Dim topBody As String
    Dim logged As Long
    Dim chunkStart As Long
    Dim chunkEnd As Long
    Dim chunk As String
    Dim cFrom As String
    Dim cFromEmail As String
    Dim msgDate As String
    Dim cSubject As String
    Dim cBody As String
    Dim cRes As String
    Dim nlPos As Long
    Dim fLine As String
    Dim eStart As Long
    Dim eEnd As Long
    Dim sPos As Long
    Dim sEnd2 As Long
    Dim sLine As String
    Dim sjPos As Long
    Dim sjEnd As Long
    Dim bStart As Long
    Dim fromName As String
    Dim fromEmail As String
    Dim toName As String
    Dim toEmail As String
    Dim isOut As Boolean
    Dim dirTxt As String
    Dim confirmMsg As String
    Dim result As String
    Dim testDate As Date
    Dim commaPos As Long

    On Error GoTo BtnErr
    
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Select an email first!", vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    
    ' Multiple selected? Log them all
    If Application.ActiveExplorer.Selection.Count > 1 Then
        Dim mc As New Collection
        Dim ms As String
        ms = ""
        Dim mx As Long
        For mx = 1 To Application.ActiveExplorer.Selection.Count
            If TypeOf Application.ActiveExplorer.Selection.Item(mx) Is Outlook.MailItem Then
                mc.Add Application.ActiveExplorer.Selection.Item(mx)
                If ms = "" Then ms = DetectShip(SafeStr(Application.ActiveExplorer.Selection.Item(mx).Subject) & " " & GetSMTPAddress(Application.ActiveExplorer.Selection.Item(mx), True))
            End If
        Next mx
        If ms = "" Then ms = InputBox("Ship?", "Ship Tracker", "DDG")
        If ms <> "" Then LogMailCollection mc, ms
        Exit Sub
    End If
    
    If Not TypeOf Application.ActiveExplorer.Selection.Item(1) Is Outlook.MailItem Then
        MsgBox "Please select an email.", vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    
    Set mail = Application.ActiveExplorer.Selection.Item(1)
    rawBody = mail.body
    bodyLen = Len(rawBody)
    
    ' Detect ship
    shipFE = GetSMTPAddress(mail, True)
    shipTE = ""
    If mail.Recipients.Count > 0 Then shipTE = GetRecipientSMTP(mail.Recipients.Item(1))
    shipName = DetectShip(SafeStr(mail.Subject) & " " & Left(rawBody, 500) & " " & SafeStr(mail.SenderName) & " " & shipFE & " " & shipTE)
    If shipName = "" Then
        shipName = InputBox("Ship not detected. Enter ship name:", "Ship Tracker", "DDG")
        If shipName = "" Then Exit Sub
    End If
    
    ' Check for thread — find From:/Sent: pairs
    ReDim fromPositions(1 To 50)
    fromCount = 0
    pos = 1
    Do
        pos = InStr(pos, rawBody, "From:", vbTextCompare)
        If pos = 0 Then Exit Do
        fromCount = fromCount + 1
        If fromCount <= 50 Then fromPositions(fromCount) = pos
        pos = pos + 5
    Loop
    
    ReDim sentPositions(1 To 50)
    sentCount = 0
    pos = 1
    Do
        pos = InStr(pos, rawBody, "Sent:", vbTextCompare)
        If pos = 0 Then Exit Do
        sentCount = sentCount + 1
        If sentCount <= 50 Then sentPositions(sentCount) = pos
        pos = pos + 5
    Loop
    
    ReDim pairPositions(1 To 50)
    pairCount = 0
    For fi = 1 To fromCount
        fp = fromPositions(fi)
        For si = 1 To sentCount
            If sentPositions(si) > fp And sentPositions(si) - fp < 500 Then
                pairCount = pairCount + 1
                If pairCount <= 50 Then pairPositions(pairCount) = fp
                Exit For
            End If
        Next si
    Next fi
    
    ' === SINGLE EMAIL (no thread) ===
    If pairCount = 0 Then
        fromName = GetRealName(SafeStr(mail.SenderName), shipFE)
        fromEmail = shipFE
        toName = ""
        toEmail = ""
        If mail.Recipients.Count > 0 Then
            toEmail = GetRecipientSMTP(mail.Recipients.Item(1))
            toName = GetRealName(SafeStr(mail.Recipients.Item(1).Name), toEmail)
        End If
        isOut = IsTeamEmail(fromEmail)
        If isOut Then dirTxt = "-> Sent by " & fromName Else dirTxt = "<- Reply from " & fromName
        
        confirmMsg = "SHIP: " & shipName & vbCrLf & _
                     "DIRECTION: " & dirTxt & vbCrLf & _
                     "FROM: " & fromName & vbCrLf & _
                     "TO: " & toName & vbCrLf & _
                     "SUBJECT: " & Left(SafeStr(mail.Subject), 60) & vbCrLf & _
                     "DATE: " & Format(mail.ReceivedTime, "m/d/yyyy h:nn AM/PM") & vbCrLf & vbCrLf & _
                     "Log this email?"
        If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Ship Tracker - " & shipName) = vbNo Then Exit Sub
        
        g_entryId = SafeStr(mail.EntryID)
        g_threadGroup = ""
        result = SendToFirebase(SafeStr(mail.Subject), CleanEmailBody(mail.body), _
            fromName, fromEmail, toName, toEmail, _
            SafeStr(mail.ConversationID), Format(mail.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z", shipName)
        
        If result = "OK" Then
            ' Mark as logged
            On Error Resume Next
            Dim manProp As Outlook.UserProperty
            Set manProp = mail.UserProperties.Add("ShipTrackerLogged", olText)
            manProp.Value = Format(Now, "yyyy-mm-ddThh:nn:ss")
            mail.Save
            On Error GoTo BtnErr
            MsgBox "Logged to " & shipName & "!" & vbCrLf & dirTxt, vbInformation, "Ship Tracker"
        Else
            MsgBox "Failed: " & result, vbCritical, "Ship Tracker"
        End If
        Exit Sub
    End If
    
    ' === THREAD DETECTED ===
    pairInfo = ""
    For fi = 1 To pairCount
        snippet = Mid(rawBody, pairPositions(fi), 50)
        snippet = Replace(Replace(Replace(snippet, vbCrLf, " "), vbCr, " "), vbLf, " ")
        pairInfo = pairInfo & fi & ". " & Left(snippet, 45) & vbCrLf
    Next fi
    
    confirmMsg = "Thread detected! " & (pairCount + 1) & " messages found." & vbCrLf & _
                 "Ship: " & shipName & vbCrLf & vbCrLf & _
                 pairInfo & vbCrLf & _
                 "Log all " & (pairCount + 1) & " messages?"
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Ship Tracker Thread - " & shipName) = vbNo Then Exit Sub
    
    ' Log top message
    g_threadGroup = "TG_" & Format(Now, "yyyymmddhhnnss") & "_" & Left(SafeStr(mail.EntryID), 10)
    topBody = Left(rawBody, pairPositions(1) - 1)
    topBody = CleanChunk(topBody)
    g_entryId = SafeStr(mail.EntryID) & "_T0"
    result = SendToFirebase(SafeStr(mail.Subject), topBody, _
        GetRealName(SafeStr(mail.SenderName), shipFE), shipFE, "", "", _
        SafeStr(mail.ConversationID), Format(mail.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z", shipName)
    logged = 0
    If result = "OK" Then logged = 1
    
    ' Log each thread message
    For ci = 1 To pairCount
        chunkStart = pairPositions(ci)
        If ci < pairCount Then
            chunkEnd = pairPositions(ci + 1) - 1
        Else
            chunkEnd = bodyLen
        End If
        chunk = Mid(rawBody, chunkStart, chunkEnd - chunkStart + 1)
        
        ' Parse From
        cFrom = ""
        cFromEmail = ""
        nlPos = InStr(chunk, vbLf)
        If nlPos = 0 Then nlPos = InStr(chunk, vbCr)
        If nlPos = 0 Then nlPos = 200
        fLine = Left(chunk, nlPos)
        fLine = Mid(fLine, 6)
        eStart = InStr(fLine, "<")
        If eStart > 0 Then
            eEnd = InStr(fLine, ">")
            If eEnd > eStart Then cFromEmail = Mid(fLine, eStart + 1, eEnd - eStart - 1)
            fLine = Left(fLine, eStart - 1)
        End If
        cFrom = GetRealName(Trim(fLine), cFromEmail)
        
        ' Parse date — handle "DayName, Month Day, Year Time AM/PM"
        msgDate = Format(mail.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z"
        sPos = InStr(1, chunk, "Sent:", vbTextCompare)
        If sPos > 0 Then
            sEnd2 = InStr(sPos, chunk, vbLf)
            If sEnd2 = 0 Then sEnd2 = InStr(sPos, chunk, vbCr)
            If sEnd2 > sPos Then
                sLine = Trim(Mid(chunk, sPos + 5, sEnd2 - sPos - 5))
                ' Strip day name: "Tuesday, August 11, 2026" -> "August 11, 2026"

                commaPos = InStr(sLine, ",")
                If commaPos > 0 And commaPos < 12 Then
                    sLine = Trim(Mid(sLine, commaPos + 1))
                End If
                If Len(sLine) > 5 Then
                    On Error Resume Next
                    testDate = CDate(sLine)
                    If Err.Number = 0 Then
                        msgDate = Format(testDate, "yyyy-mm-ddThh:nn:ss") & "Z"
                    End If
                    Err.Clear
                    On Error GoTo BtnErr
                End If
            End If
        End If
        
        ' Parse Subject
        cSubject = SafeStr(mail.Subject)
        sjPos = InStr(1, chunk, "Subject:", vbTextCompare)
        If sjPos > 0 Then
            sjEnd = InStr(sjPos, chunk, vbLf)
            If sjEnd = 0 Then sjEnd = InStr(sjPos, chunk, vbCr)
            If sjEnd > sjPos Then cSubject = Trim(Mid(chunk, sjPos + 8, sjEnd - sjPos - 8))
        End If
        
        ' Parse Body
        cBody = ""
        If sjPos > 0 Then
            bStart = InStr(sjPos, chunk, vbLf)
            If bStart = 0 Then bStart = InStr(sjPos, chunk, vbCr)
            If bStart > 0 Then cBody = CleanChunk(Mid(chunk, bStart + 1))
        End If
        
        If Len(cFrom) > 0 Then
            g_entryId = SafeStr(mail.EntryID) & "_T" & CStr(ci)
            cRes = SendToFirebase(cSubject, cBody, cFrom, cFromEmail, "", "", _
                SafeStr(mail.ConversationID), msgDate, shipName)
            If cRes = "OK" Then logged = logged + 1
        End If
        DoEvents
    Next ci
    
    MsgBox "Logged " & logged & " of " & (pairCount + 1) & " to " & shipName, vbInformation, "Ship Tracker"
    Exit Sub

BtnErr:
    MsgBox "Error: " & Err.Description & " (#" & Err.Number & ")", vbCritical, "Ship Tracker"
End Sub

Private Function CleanChunk(txt As String) As String
    Dim result As String
    result = txt
    
    ' Remove everything after signature markers
    Dim sigMarkers As Variant
    sigMarkers = Array(vbLf & "V/R", vbCr & "V/R", vbLf & "v/r", vbCr & "v/r", _
                       vbLf & "Very Respectfully", vbCr & "Very Respectfully", _
                       vbLf & "Respectfully,", vbCr & "Respectfully,", _
                       vbLf & "Thanks,", vbCr & "Thanks,", _
                       vbLf & "Thank you,", vbCr & "Thank you,")
    Dim sg As Long
    For sg = 0 To UBound(sigMarkers)
        Dim sigPos As Long
        sigPos = InStr(1, result, sigMarkers(sg), vbTextCompare)
        If sigPos > 0 Then result = Left(result, sigPos - 1)
    Next sg
    
    ' Remove lines with email addresses
    result = RemoveEmailLines(result)
    
    ' Clean whitespace
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    result = Replace(result, vbCrLf & vbCrLf & vbCrLf, vbCrLf)
    result = Replace(result, vbLf & vbLf & vbLf, vbLf)
    result = Replace(result, vbCr & vbLf, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbCr, " ")
    
    result = Trim(result)
    If Len(result) > 1500 Then result = Left(result, 1500)
    CleanChunk = result
End Function

' Remove lines that are just email addresses
Private Function RemoveEmailLines(txt As String) As String
    Dim result As String: result = ""
    Dim lines() As String
    lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim l As String: l = Trim(lines(i))
        ' Skip lines that are mainly email addresses or phone numbers
        If l Like "*<mailto:*" Then GoTo SkipLine
        If l Like "NIPR:*" Or l Like "SIPR:*" Then GoTo SkipLine
        If l Like "Desk Phone:*" Or l Like "Phone:*" Or l Like "DSN:*" Or l Like "Comm:*" Then GoTo SkipLine
        If Len(l) < 60 And InStr(l, "<") > 0 And InStr(l, "@") > 0 Then GoTo SkipLine
        If Len(l) < 20 And l Like "*###*###*####*" Then GoTo SkipLine
        If l Like "(###)###-####*" Then GoTo SkipLine
        ' Skip title/role lines in sig
        If LCase(l) Like "*technician*" And Len(l) < 40 Then GoTo SkipLine
        If LCase(l) Like "*nswc corona*" Then GoTo SkipLine
        If LCase(l) Like "*combat systems officer*" Then GoTo SkipLine
        If LCase(l) Like "*uss *" And Len(l) < 50 Then GoTo SkipLine
        If l = "" Then GoTo SkipLine
        
        result = result & l & " "
SkipLine:
    Next i
    RemoveEmailLines = Trim(result)
End Function

Private Sub LogMailCollection(col As Collection, shipName As String)
    Dim preview As String: preview = "Ship: " & shipName & " | Emails: " & col.Count & vbCrLf & vbCrLf
    Dim pi As Long
    For pi = 1 To col.Count
        Dim pm As Outlook.MailItem: Set pm = col.Item(pi)
        preview = preview & Format(pm.ReceivedTime, "m/d h:nn") & " " & GetRealName(SafeStr(pm.SenderName), GetSMTPAddress(pm, True)) & vbCrLf
    Next pi
    If MsgBox(preview & vbCrLf & "Log all?", vbYesNo, "Ship Tracker") = vbNo Then Exit Sub
    Dim logged As Long: logged = 0
    For pi = 1 To col.Count
        Set pm = col.Item(pi)
        Dim fe As String: fe = GetSMTPAddress(pm, True)
        Dim tn As String: tn = "": Dim te As String: te = ""
        If pm.Recipients.Count > 0 Then tn = GetRealName(SafeStr(pm.Recipients.Item(1).Name), GetRecipientSMTP(pm.Recipients.Item(1))): te = GetRecipientSMTP(pm.Recipients.Item(1))
        g_entryId = SafeStr(pm.EntryID)
        If SendToFirebase(SafeStr(pm.Subject), CleanEmailBody(pm.body), GetRealName(SafeStr(pm.SenderName), fe), fe, tn, te, SafeStr(pm.ConversationID), Format(pm.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z", shipName) = "OK" Then logged = logged + 1
        DoEvents
    Next pi
    MsgBox "Logged " & logged & " of " & col.Count, vbInformation, "Ship Tracker"
End Sub

Private Function RemoveReFw(subj As String) As String
    Dim s As String: s = Trim(subj)
    Do While UCase(Left(s, 3)) = "RE:" Or UCase(Left(s, 3)) = "FW:" Or UCase(Left(s, 4)) = "FWD:"
        If UCase(Left(s, 4)) = "FWD:" Then s = Trim(Mid(s, 5)) Else s = Trim(Mid(s, 4))
    Loop
    RemoveReFw = s
End Function

Private Function MinVal(a As Long, b As Long) As Long
    If a < b Then MinVal = a Else MinVal = b
End Function

Private Function MaxVal(a As Long, b As Long) As Long
    If a > b Then MaxVal = a Else MaxVal = b
End Function

' Decode Base64 string to text
Private Function DecodeBase64(b64 As String) As String
    On Error Resume Next
    If Len(b64) = 0 Then DecodeBase64 = "": Exit Function
    
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write node.nodeTypedValue
    stream.Position = 0
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    DecodeBase64 = stream.ReadText
    stream.Close
    
    Set stream = Nothing
    Set node = Nothing
    Set xml = Nothing
End Function

Private Function IsSignatureLine(line As String) As Boolean
    Dim l As String: l = LCase(Trim(line))
    IsSignatureLine = (l = "v/r" Or l = "v/r," Or l = "vr" Or l = "vr," Or l = "respectfully," Or l = "very respectfully," Or l = "regards," Or l = "best regards," Or l = "kind regards," Or l = "thanks," Or l = "thank you," Or l = "thanks" Or l = "thank you" Or l = "sincerely," Or Left(l, 10) = "sent from " Or Left(l, 4) = "tel:" Or Left(l, 6) = "phone:" Or Left(l, 4) = "dsn:" Or Left(l, 5) = "comm:" Or Left(l, 7) = "mobile:" Or (Len(l) < 20 And l Like "*###-###-####*"))
End Function

Public Sub LogSelectedToShipTracker()
    On Error Resume Next
    
    Dim olExplorer As Outlook.Explorer
    Set olExplorer = Application.ActiveExplorer
    
    If olExplorer.Selection.Count = 0 Then
        MsgBox "Select one or more emails!", vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    
    Dim cnt As Long: cnt = olExplorer.Selection.Count
    If MsgBox("Log " & cnt & " email(s)?", vbYesNo + vbQuestion, "Ship Tracker Bulk") = vbNo Then Exit Sub
    
    Dim logged As Long, failed As Long, skipped As Long
    Dim i As Long
    
    For i = 1 To cnt
        If TypeOf olExplorer.Selection.Item(i) Is Outlook.MailItem Then
            Dim m As Outlook.MailItem
            Set m = olExplorer.Selection.Item(i)
            
                Dim fe As String: fe = GetSMTPAddress(m, True)
                Dim tn As String: tn = ""
                Dim te As String: te = ""
                If m.Recipients.Count > 0 Then
                    tn = SafeStr(m.Recipients.Item(1).Name)
                    te = GetRecipientSMTP(m.Recipients.Item(1))
                End If
            
            Dim sn As String
            sn = DetectShip(SafeStr(m.subject) & " " & SafeStr(Left(m.body, 500)) & " " & SafeStr(m.SenderName) & " " & fe & " " & tn & " " & te)
            
            ' Send ALL emails — even without ship match. App inbox handles assignment.
            If sn = "" Then sn = "UNMATCHED"
            
                Dim res As String
                res = SendToFirebase(SafeStr(m.subject), CleanEmailBody(m.body), _
                    GetRealName(SafeStr(m.SenderName), fe), fe, tn, te, _
                    SafeStr(m.ConversationID), Format(m.ReceivedTime, "yyyy-mm-ddThh:nn:ss") & "Z", sn)
                
                If res = "OK" Then
                    If sn = "UNMATCHED" Then
                        skipped = skipped + 1
                    Else
                        logged = logged + 1
                    End If
                Else
                    failed = failed + 1
                End If
        End If
    Next i
    
    MsgBox "Done!" & vbCrLf & vbCrLf & _
           "Logged: " & logged & vbCrLf & _
           "Failed: " & failed & vbCrLf & _
           "Skipped (no ship): " & skipped, vbInformation, "Ship Tracker Bulk"
End Sub

' ============================================
' SEND TO FIREBASE — returns "OK" or error msg
' ============================================
Private Function SendToFirebase(subject As String, body As String, _
    fromName As String, fromEmail As String, toName As String, toEmail As String, _
    convId As String, recvAt As String, shipName As String) As String
    
    On Error GoTo FBError
    
    Dim j As String
    j = "{""fields"": {" & _
        SF("subject", subject) & "," & _
        SF("body", body) & "," & _
        SF("senderName", fromName) & "," & _
        SF("senderEmail", fromEmail) & "," & _
        SF("from", fromName & " <" & fromEmail & ">") & "," & _
        SF("toName", toName) & "," & _
        SF("toEmail", toEmail) & "," & _
        SF("to", toName & " <" & toEmail & ">") & "," & _
        SF("conversationId", convId) & "," & _
        SF("receivedAt", recvAt) & "," & _
        SF("importedBy", MY_NAME) & "," & _
        SF("flowOwner", MY_NAME) & "," & _
        SF("matchedShip", shipName) & "," & _
        SF("matchMethod", "vba") & "," & _
        SF("outlookEntryId", g_entryId) & "," & _
        SF("threadGroup", g_threadGroup) & "," & _
        """processed"": {""booleanValue"": false}," & _
        SF("timestamp", Format(Now, "yyyy-mm-ddThh:nn:ss") & "Z") & _
        "}}"
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", FIREBASE_URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.Send j
    
    If http.Status >= 200 And http.Status < 300 Then
        SendToFirebase = "OK"
    Else
        SendToFirebase = "HTTP " & http.Status & ": " & Left(http.responseText, 200)
    End If
    
    Set http = Nothing
    Exit Function
    
FBError:
    SendToFirebase = "VBA Error: " & Err.Description
End Function

' ============================================
' HELPERS
' ============================================

' Build a JSON string field
Private Function SF(key As String, val As String) As String
    SF = """" & key & """: {""stringValue"": """ & CleanJ(val) & """}"
End Function

' Clean string for JSON — handles ALL problem characters
Private Function CleanJ(s As String) As String
    Dim r As String: r = s
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, " ")
    r = Replace(r, vbCr, " ")
    r = Replace(r, vbLf, " ")
    r = Replace(r, vbTab, " ")
    ' Remove ALL control characters (0-31) and DEL (127)
    Dim i As Long
    For i = 0 To 31
        r = Replace(r, Chr(i), "")
    Next i
    r = Replace(r, Chr(127), "")
    ' Remove non-ASCII that could break JSON
    Dim result As String: result = ""
    For i = 1 To Len(r)
        Dim c As Long: c = AscW(Mid(r, i, 1))
        If c >= 32 And c < 127 Then
            result = result & Mid(r, i, 1)
        ElseIf c = 8217 Or c = 8216 Then
            result = result & "'"
        ElseIf c = 8220 Or c = 8221 Then
            result = result & """"
        ElseIf c = 8211 Or c = 8212 Then
            result = result & "-"
        ElseIf c = 160 Then
            result = result & " "
        ElseIf c > 127 Then
            result = result & " "
        End If
    Next i
    CleanJ = result
End Function

' Safe string — handles Nothing and nulls
Private Function SafeStr(s As Variant) As String
    On Error Resume Next
    If IsNull(s) Or IsEmpty(s) Then SafeStr = "" Else SafeStr = CStr(s)
    On Error GoTo 0
End Function

' Get SMTP email from sender (handles Exchange)
Private Function GetSMTPAddress(mail As Outlook.MailItem, isSender As Boolean) As String
    On Error Resume Next
    Dim addr As String
    
    If isSender Then
        addr = mail.SenderEmailAddress
        If Left(addr, 1) = "/" Or InStr(addr, "@") = 0 Then
            If Not mail.sender Is Nothing Then
                If mail.sender.AddressEntryUserType = olExchangeUserAddressEntry Or _
                   mail.sender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                    addr = mail.sender.GetExchangeUser.PrimarySmtpAddress
                End If
            End If
        End If
    End If
    
    If addr = "" Then addr = "unknown"
    GetSMTPAddress = SafeStr(addr)
    On Error GoTo 0
End Function

' Get SMTP from recipient
Private Function GetRecipientSMTP(recip As Outlook.Recipient) As String
    On Error Resume Next
    Dim addr As String
    addr = recip.Address
    If Left(addr, 1) = "/" Or InStr(addr, "@") = 0 Then
        If recip.AddressEntry.AddressEntryUserType = olExchangeUserAddressEntry Or _
           recip.AddressEntry.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            addr = recip.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        End If
    End If
    If addr = "" Then addr = "unknown"
    GetRecipientSMTP = SafeStr(addr)
    On Error GoTo 0
End Function

' Detect ship from text
Private Function DetectShip(text As String) As String
    Dim ships As Variant
    ' Use config ships if available
    If g_configLoaded And Len(g_configShips) > 5 Then
        ships = Split(g_configShips, ",")
    Else
        ships = Array("DDG54", "DDG55", "DDG56", "DDG57", "DDG62", "DDG63", "DDG67", "DDG70", _
                      "DDG71", "DDG72", "DDG77", "DDG82", "DDG96", "DDG98", "DDG101", "DDG103", _
                      "DDG109", "DDG111", _
                      "CG70", "CG52", "CG53", "CG54", "CG55", "CG56", "CG57", "CG58", "CG59", "CG60", _
                      "CG61", "CG62", "CG63", "CG64", "CG65", "CG66", "CG67", "CG68", "CG69", "CG71", "CG72", "CG73", _
                      "LCS1", "LCS2", "LCS3", "LCS4", "LCS5", "LCS6", "LCS7", "LCS8", "LCS9", "LCS10", _
                      "LPD17", "LPD18", "LPD19", "LPD20", "LPD21", "LPD22", "LPD23", "LPD24", "LPD25", "LPD26", _
                      "LHD1", "LHD2", "LHD3", "LHD4", "LHD5", "LHD6", "LHD7", "LHD8", _
                      "CVN68", "CVN69", "CVN70", "CVN71", "CVN72", "CVN73", "CVN74", "CVN75", "CVN76", "CVN77", "CVN78", "CVN79")
    End If
    Dim st As String: st = UCase(text)
    Dim i As Long
    For i = LBound(ships) To UBound(ships)
        If InStr(1, st, ships(i)) > 0 Then DetectShip = ships(i): Exit Function
    Next i
    For i = LBound(ships) To UBound(ships)
        Dim num As String: num = Mid(ships(i), 4)
        Dim prefix As String: prefix = Left(ships(i), 3)
        If Left(prefix, 2) = "CG" Then prefix = "CG"
        If Left(prefix, 3) = "LCS" Then prefix = "LCS"
        If Left(prefix, 3) = "LPD" Then prefix = "LPD"
        If Left(prefix, 3) = "LHD" Then prefix = "LHD"
        If Left(prefix, 3) = "CVN" Then prefix = "CVN"
        If InStr(1, st, prefix & " " & num) > 0 Or InStr(1, st, prefix & "-" & num) > 0 Then
            DetectShip = ships(i): Exit Function
        End If
    Next i
    ' Check for ship NAMES (not just hull numbers)
    Dim shipNames As Variant
    shipNames = Array( _
        "CURTIS WILBUR,DDG54", "STOUT,DDG55", "JOHN S MCCAIN,DDG56", "MITSCHER,DDG57", _
        "FITZGERALD,DDG62", "STETHEM,DDG63", "COLE,DDG67", "HOPPER,DDG70", _
        "ROSS,DDG71", "MAHAN,DDG72", "O'KANE,DDG77", "LASSEN,DDG82", _
        "BAINBRIDGE,DDG96", "FORREST SHERMAN,DDG98", "GRIDLEY,DDG101", "TRUXTUN,DDG103", _
        "JASON DUNHAM,DDG109", "SPRUANCE,DDG111", _
        "LAKE ERIE,CG70", "BUNKER HILL,CG52", "MOBILE BAY,CG53", _
        "ANTIETAM,CG54", "LEYTE GULF,CG55", "SAN JACINTO,CG56", _
        "LAKE CHAMPLAIN,CG57", "PHILIPPINE SEA,CG58", "PRINCETON,CG59", _
        "NORMANDY,CG60", "MONTEREY,CG61", "CHANCELLORSVILLE,CG62", _
        "COWPENS,CG63", "GETTYSBURG,CG64", "CHOSIN,CG65", _
        "HUE CITY,CG66", "SHILOH,CG67", "ANZIO,CG68", _
        "VICKSBURG,CG69", "CAPE ST GEORGE,CG71", "VELLA GULF,CG72", "PORT ROYAL,CG73")
    
    Dim sn As Long
    For sn = LBound(shipNames) To UBound(shipNames)
        Dim nameParts() As String
        nameParts = Split(shipNames(sn), ",")
        If UBound(nameParts) >= 1 Then
            If InStr(1, st, UCase(nameParts(0))) > 0 Then
                DetectShip = nameParts(1)
                Exit Function
            End If
        End If
    Next sn
    
    DetectShip = ""
End Function

' Clean email body — extract only the actual message content
Private Function CleanEmailBody(rawBody As String) As String
    Dim lines() As String
    Dim result As String
    Dim i As Long
    Dim line As String
    Dim lineLower As String
    
    result = ""
    lines = Split(rawBody, vbCrLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        lineLower = LCase(line)
        
        ' Stop at forwarding/reply headers
        If Left(line, 5) = "From:" And InStr(line, "@") > 0 Then Exit For
        If Left(line, 5) = "Sent:" And (InStr(line, "AM") > 0 Or InStr(line, "PM") > 0) Then Exit For
        If Left(line, 3) = "To:" And InStr(line, "@") > 0 Then Exit For
        If Left(line, 3) = "Cc:" And InStr(line, "@") > 0 Then Exit For
        If line = "________________________________" Then Exit For
        If Left(line, 22) = "-----Original Message-" Then Exit For
        If Left(line, 10) = "----------" And InStr(line, "Forwarded") > 0 Then Exit For
        
        ' Stop at signature blocks
        If line = "-- " Or line = "--" Then Exit For
        If lineLower = "v/r" Or lineLower = "v/r," Then Exit For
        If lineLower = "vr" Or lineLower = "vr," Then Exit For
        If lineLower = "respectfully," Or lineLower = "respectfully" Then Exit For
        If lineLower = "very respectfully," Or lineLower = "very respectfully" Then Exit For
        If lineLower = "regards," Or lineLower = "regards" Then Exit For
        If lineLower = "best regards," Or lineLower = "best regards" Then Exit For
        If lineLower = "warm regards," Or lineLower = "warm regards" Then Exit For
        If lineLower = "kind regards," Or lineLower = "kind regards" Then Exit For
        If lineLower = "best," Or lineLower = "best" Then Exit For
        If lineLower = "thank you," Or lineLower = "thanks," Then Exit For
        If lineLower = "thank you" Or lineLower = "thanks" Then Exit For
        If lineLower = "sincerely," Or lineLower = "sincerely" Then Exit For
        If Left(lineLower, 6) = "cheers" Then Exit For
        If Left(lineLower, 10) = "sent from " Then Exit For
        If Left(lineLower, 13) = "get outlook f" Then Exit For
        ' Additional signature patterns
        If lineLower = "warm regards," Or lineLower = "warm regards" Then Exit For
        If lineLower = "best wishes," Or lineLower = "best wishes" Then Exit For
        If Left(lineLower, 11) = "classfied b" Then Exit For
        If Left(lineLower, 14) = "classification:" Then Exit For
        If lineLower Like "*nswc corona*" Then Exit For
        If lineLower Like "*information system*" And Len(line) < 50 Then Exit For
        If lineLower Like "*combat systems*" And Len(line) < 50 Then Exit For
        If lineLower Like "*system test officer*" Then Exit For
        
        ' Stop at lines that look like signature info (title, phone, address)
        If Left(lineLower, 4) = "tel:" Or Left(lineLower, 6) = "phone:" Or Left(lineLower, 4) = "dsn:" Then Exit For
        If Left(lineLower, 5) = "comm:" Or Left(lineLower, 7) = "mobile:" Or Left(lineLower, 5) = "cell:" Then Exit For
        If Left(lineLower, 8) = "sipr: " Or Left(lineLower, 8) = "nipr: " Then Exit For
        
        ' Skip lines that are just email addresses
        If InStr(line, "<") > 0 And InStr(line, "@") > 0 And InStr(line, ">") > 0 And Len(line) < 100 Then
            ' Address line — skip
        ' Skip lines that are just a phone number
        ElseIf Len(line) < 20 And line Like "*###-###-####*" Then
            ' Phone number — skip
        ' Skip lines that are just a rank/name/title block
        ElseIf Len(line) < 50 And (InStr(lineLower, "information system") > 0 Or InStr(lineLower, "technician") > 0 Or InStr(lineLower, "specialist") > 0) And InStr(line, " ") > 0 Then
            ' Title line — likely start of sig
            Exit For
        ElseIf line <> "" Then
            If result <> "" Then result = result & " "
            result = result & line
        End If
    Next i
    
    ' Final cleanup — remove trailing partial sentences from sig cutoff
    ' If result ends with a name-like pattern after our content, trim it
    Dim lastPeriod As Long
    lastPeriod = InStrRev(result, ".")
    If lastPeriod > 0 And Len(result) - lastPeriod > 50 Then
        ' More than 50 chars after last period — might be sig remnant
        Dim afterPeriod As String
        afterPeriod = Mid(result, lastPeriod + 1)
        If InStr(LCase(afterPeriod), "@") > 0 Or InStr(LCase(afterPeriod), "dsn") > 0 Or InStr(LCase(afterPeriod), "phone") > 0 Then
            result = Left(result, lastPeriod)
        End If
    End If
    
    ' Cap at 1500 chars
    If Len(result) > 1500 Then result = Left(result, 1500)
    
    CleanEmailBody = SafeStr(result)
End Function

' Check if from team
' Check if email belongs to a REAP team member specifically
Private Function IsTeamMember(email As String) As Boolean
    Dim a As String: a = LCase(email)
    IsTeamMember = InStr(a, "daniel.k.brown") > 0 Or InStr(a, "dbrown") > 0 Or _
                   InStr(a, "clyde.a.grant") > 0 Or InStr(a, "cgrant") > 0 Or _
                   InStr(a, "uriel.pena") > 0 Or InStr(a, "up_reap") > 0 Or _
                   InStr(a, "kaazim.j.hanif") > 0 Or InStr(a, "jayreap") > 0 Or _
                   InStr(a, "damian.j.pierman") > 0 Or InStr(a, "eryk.kedzierski") > 0 Or _
                   InStr(a, "steven.m.spellman") > 0 Or InStr(a, "mohammed.cardenas") > 0
    ' Also check config team roster
    If Not IsTeamMember And g_configLoaded And Len(g_configTeam) > 5 Then
        Dim teamLines() As String
        teamLines = Split(Replace(Replace(g_configTeam, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        Dim tl As Long
        For tl = 0 To UBound(teamLines)
            If InStr(teamLines(tl), "=") > 0 Then
                Dim matchKey As String: matchKey = LCase(Trim(Split(teamLines(tl), "=")(0)))
                If Len(matchKey) > 0 And InStr(a, matchKey) > 0 Then IsTeamMember = True: Exit Function
            End If
        Next tl
    End If
End Function

Private Function IsTeamEmail(email As String) As Boolean
    Dim a As String: a = LCase(email)
    IsTeamEmail = (InStr(a, "navy.mil") > 0 And InStr(a, "ddg") = 0 And InStr(a, "@cg") = 0 And InStr(a, "@lcs") = 0 And InStr(a, "@lpd") = 0 And InStr(a, "@lhd") = 0 And InStr(a, "@cvn") = 0) Or _
                  InStr(a, "sonar.navy") > 0 Or InStr(a, "niwc.navy") > 0 Or _
                  InStr(a, "spawar.navy") > 0 Or InStr(a, "mcas") > 0
End Function

' ============================================
' END OF DAY EMAIL - one click builds the finished EOD email
' Ship Tracker keeps today's report + Excel saved while Daily Checks is open;
' this pulls it down and opens a ready-to-send email with the Excel attached.
'
' ADD THE BUTTON (once):
'   File > Options > Quick Access Toolbar > Choose commands from: Macros
'   > Project1.ThisOutlookSession.CreateEodEmail > Add > OK
' FIRST USE (once per PC):
'   Click the button and sign in with your Ship Tracker username and password.
' ============================================

Public Sub CreateEodEmail()
    On Error GoTo Failed

    Dim idToken As String
    idToken = GetShipTrackerToken()
    If idToken = "" Then Exit Sub

    Dim dayKey As String: dayKey = Format(Date, "yyyy-mm-dd")
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", EOD_DOC_URL & dayKey, False
    http.setRequestHeader "Authorization", "Bearer " & idToken
    http.Send

    If http.Status = 404 Then
        MsgBox "There's no End of Day report for today yet." & vbCrLf & vbCrLf & _
               "Open Daily Checks in Ship Tracker for a few seconds (it saves the report automatically), then click EOD Email again.", _
               vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    If http.Status = 401 Or http.Status = 403 Then
        DeleteShipTrackerLink
        MsgBox "Ship Tracker didn't accept this login (your account may not be approved)." & vbCrLf & vbCrLf & _
               "Click EOD Email to sign in again.", _
               vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    If http.Status <> 200 Then
        MsgBox "Couldn't download today's report (HTTP " & http.Status & ").", vbCritical, "Ship Tracker"
        Exit Sub
    End If

    Dim resp As String: resp = http.responseText
    Set http = Nothing

    Dim subject As String: subject = ExtractFirestoreField(resp, "emailSubject")
    Dim reportHtml As String: reportHtml = DecodeBase64(ExtractFirestoreField(resp, "emailHtmlB64"))
    Dim xlsxB64 As String: xlsxB64 = ExtractFirestoreField(resp, "xlsxB64")

    If reportHtml = "" Then
        MsgBox "Today's report hasn't been prepared for Outlook yet." & vbCrLf & vbCrLf & _
               "Open Daily Checks in Ship Tracker for a few seconds, then click EOD Email again.", _
               vbExclamation, "Ship Tracker"
        Exit Sub
    End If
    If subject = "" Then subject = "REAP Status Report - " & Format(Date, "mmmm d, yyyy")

    ' Build the email; Display first so Outlook adds the user's signature, then put the report above it
    Dim mail As Outlook.MailItem
    Set mail = Application.CreateItem(0)   ' olMailItem
    mail.Subject = subject
    mail.Display

    Dim body As String: body = mail.HTMLBody
    Dim bodyTag As Long: bodyTag = InStr(1, LCase(body), "<body")
    If bodyTag > 0 Then
        Dim tagEnd As Long: tagEnd = InStr(bodyTag, body, ">")
        mail.HTMLBody = Left(body, tagEnd) & reportHtml & "<p>&nbsp;</p>" & Mid(body, tagEnd + 1)
    Else
        mail.HTMLBody = reportHtml & body
    End If

    ' Attach the Excel sheet
    If xlsxB64 <> "" Then
        Dim xlsxPath As String
        xlsxPath = Environ("TEMP") & "\REAP_Status_Report_" & dayKey & ".xlsx"
        If SaveBase64ToFile(xlsxB64, xlsxPath) Then
            mail.Attachments.Add xlsxPath
        Else
            MsgBox "The email is ready, but the Excel file couldn't be attached.", vbExclamation, "Ship Tracker"
        End If
    End If
    Exit Sub

Failed:
    MsgBox "EOD Email failed: " & Err.Description, vbCritical, "Ship Tracker"
End Sub

' Returns a fresh sign-in token for Ship Tracker, asking for the login the first time ("" = cancelled/failed)
Private Function GetShipTrackerToken() As String
    On Error Resume Next
    GetShipTrackerToken = ""
    Dim http As Object
    Dim resp As String

    Dim refreshToken As String
    refreshToken = GetSetting("ShipTracker", "Auth", "RefreshToken", "")

    If refreshToken <> "" Then
        ' Normal case: renew the login saved on this PC
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
        http.Open "POST", "https://securetoken.googleapis.com/v1/token?key=" & FIREBASE_WEB_KEY, False
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.Send "grant_type=refresh_token&refresh_token=" & refreshToken
        If http.Status = 200 Then
            resp = http.responseText
            Dim renewed As String: renewed = ExtractJsonString(resp, "refresh_token")
            If renewed <> "" Then SaveSetting "ShipTracker", "Auth", "RefreshToken", renewed
            GetShipTrackerToken = ExtractJsonString(resp, "id_token")
            Exit Function
        End If
        ' Saved login stopped working (password changed, account removed...) - sign in again below
        DeleteShipTrackerLink
    End If

    ' First time on this PC: sign in with the Ship Tracker username + password
    Dim userName As String
    userName = Trim(InputBox("Sign in to Ship Tracker (one time on this PC)." & vbCrLf & vbCrLf & _
               "Ship Tracker username:", "Ship Tracker"))
    If userName = "" Then Exit Function
    Dim pw As String
    pw = InputBox("Ship Tracker password for " & userName & ":" & vbCrLf & vbCrLf & _
         "Outlook can't hide what you type in this box. Your password goes only to the Ship Tracker sign-in and is never saved.", _
         "Ship Tracker")
    If pw = "" Then Exit Function

    Dim body As String
    body = "{""email"":""" & JsonEscape(LCase(userName) & "@ship-checker.firebaseapp.com") & _
           """,""password"":""" & JsonEscape(pw) & """,""returnSecureToken"":true}"
    pw = ""
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=" & FIREBASE_WEB_KEY, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send body
    body = ""

    resp = http.responseText
    If http.Status <> 200 Then
        If InStr(resp, "INVALID_LOGIN_CREDENTIALS") > 0 Or InStr(resp, "INVALID_PASSWORD") > 0 Or _
           InStr(resp, "EMAIL_NOT_FOUND") > 0 Or InStr(resp, "INVALID_EMAIL") > 0 Then
            MsgBox "Wrong Ship Tracker username or password. Click EOD Email to try again.", vbExclamation, "Ship Tracker"
        ElseIf InStr(resp, "TOO_MANY_ATTEMPTS") > 0 Then
            MsgBox "Too many sign-in attempts. Wait a few minutes, then try again.", vbExclamation, "Ship Tracker"
        Else
            MsgBox "Couldn't sign in to Ship Tracker (HTTP " & http.Status & ").", vbExclamation, "Ship Tracker"
        End If
        Exit Function
    End If

    Dim saved As String: saved = ExtractJsonString(resp, "refreshToken")
    If saved <> "" Then SaveSetting "ShipTracker", "Auth", "RefreshToken", saved
    GetShipTrackerToken = ExtractJsonString(resp, "idToken")
End Function

' Escape a value for a JSON string
Private Function JsonEscape(s As String) As String
    JsonEscape = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

Private Sub DeleteShipTrackerLink()
    On Error Resume Next
    DeleteSetting "ShipTracker", "Auth", "RefreshToken"
End Sub

' Value of a top-level "key": "value" pair in plain JSON (tokens have no quotes or escapes)
Private Function ExtractJsonString(json As String, key As String) As String
    Dim p As Long: p = InStr(1, json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, json, ":")
    If p = 0 Then Exit Function
    Dim q1 As Long: q1 = InStr(p, json, """")
    If q1 = 0 Then Exit Function
    Dim q2 As Long: q2 = InStr(q1 + 1, json, """")
    If q2 = 0 Then Exit Function
    ExtractJsonString = Mid(json, q1 + 1, q2 - q1 - 1)
End Function

' Decode base64 straight to a binary file (for the .xlsx attachment)
Private Function SaveBase64ToFile(b64 As String, path As String) As Boolean
    On Error GoTo Failed
    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object: Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile path, 2 ' overwrite
    stream.Close
    SaveBase64ToFile = True
    Exit Function
Failed:
    SaveBase64ToFile = False
End Function
