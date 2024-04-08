VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ExtraRequestsCtrl 
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   ScaleHeight     =   9570
   ScaleWidth      =   17115
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   12000
      Picture         =   "ExtraRequestsCtrl.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   14160
      Top             =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Removal from Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdReport 
         Caption         =   "Report Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2760
         Picture         =   "ExtraRequestsCtrl.ctx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEntityBarcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblMicrotom 
         Caption         =   "Microtome: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode Entity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presentation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   4680
         Picture         =   "ExtraRequestsCtrl.ctx":058C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdShowRequests 
         Caption         =   "Show    Requests"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   3000
         Picture         =   "ExtraRequestsCtrl.ctx":09CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstEntityTypes 
         Height          =   900
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label Label1 
         Caption         =   "Entity Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   12726
      _Version        =   393216
   End
End
Attribute VB_Name = "ExtraRequestsCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Option Explicit


Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser


Private con As New ADODB.Connection
Private entityType As String

'holds the entities of the presented list
'key  - the entity name
'item - a collection of locations on the grid for this entity
Private dicEntities As New Dictionary

'holds the extra_request_data_id against the entity name:
Private dicBarcodeEntities As New Dictionary

Private Const MARK_SELECTED = &HC0FFFF

'used for the printing of the
'grid to the printer:
Private RowSize As Long
Private ColSize As Long
Private LeftOffset As Long
Private TopOffset As Long

Private Const MAX_DIGITS_PER_CELL = 24

'data of the font to be used when printing the aliquot names.
'to be read from a relevant phrase
Private strFontName As String
Private iFontSize As Integer
Private isFontBold As Boolean

Private Const MAX_LINES = 20

Private dicOperatorAllowedToReportReserveSlides As New Dictionary

Private sdg_log As New SdgLog.CreateLog


Private Sub cmdClose_Click()
10        Call NtlsSite.CloseWindow
End Sub

'print the grid contents;
'on each page print up to MAX_LINES rows of the grid;
'the column headers are printed on each page;
Private Sub cmdPrint_Click()
20    On Error GoTo ERR_cmdPrint_Click

          Dim i As Integer
          Dim iFirstLine As Integer
          Dim iLastLine As Integer
          
30        i = 1
          
40        While i < grid.Rows
50            iFirstLine = i
60            iLastLine = IIf(i + MAX_LINES >= grid.Rows, grid.Rows - 1, i + MAX_LINES - 1)

70            Call PrintToPrinter(iFirstLine, iLastLine, MAX_LINES)

80            i = i + MAX_LINES
90        Wend
          
100       lblMicrotom.Visible = False


          'for reserve slides - report the removal automatically:
      '    If lstEntityTypes.Text = "Reserve-Slide" Then
      '        Call ReportReserveSlides
      '    End If
          

110       Exit Sub
ERR_cmdPrint_Click:
120   MsgBox "ERR_cmdPrint_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'barcode all presented entities:
Private Sub BarcodeAll()
130   On Error GoTo ERR_BarcodeAll

          Dim i As Integer

140       For i = 0 To dicEntities.Count - 1
150           Call BarcodeEntity(CStr(dicEntities.Keys(i)))
160       Next i

170       Exit Sub
ERR_BarcodeAll:
180   MsgBox "ERR_BarcodeAll" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

' 1. simulate the barcode of all the entities
' 2. simulate pressing the report button
' results in changing the status of the extra request
' & changing the field U_ARCHIVE to false on the entity
'Private Sub ReportReserveSlides()
'On Error GoTo ERR_ReportReserveSlides
'    Dim i As Integer
'
'    For i = 0 To dicEntities.Count - 1
'        Call BarcodeEntity(CStr(dicEntities.Keys(i)))
'    Next i
'
'    Call cmdReport_Click
'
'    Exit Sub
'ERR_ReportReserveSlides:
'MsgBox "ERR_ReportReserveSlides" & vbCrLf &  "In Line #" & Erl & vbCrLf & Err.Description
'End Sub

'change status 'V'->'P' for all items selected by barcode;
'refresh selection of entities of the current type;
'if all entities of the current type are reported, delete this entity-type from
'the list
Private Sub cmdReport_Click()
190   On Error GoTo ERR_cmdReport_Click
          Dim i As Integer
          Dim k As Integer
          Dim sql As String
          Dim rs As Recordset
          Dim strLastEntity As String
          Dim strExternalRef As String
          Dim strSdgId As String
         
200       For i = 0 To dicBarcodeEntities.Count - 1
210           sql = " update lims_sys.u_extra_request_data_user"
220           sql = sql & " set u_status = 'P'"
230           sql = sql & " where u_extra_request_data_id = " & dicBarcodeEntities.Keys(i)
240           con.Execute (sql)
              'update sdg_log table (once per entity)
250           If strLastEntity <> CStr(dicBarcodeEntities.Items(i)) Then

260               strSdgId = GetSdgForEntity(CStr(dicBarcodeEntities.Items(i)), entityType)

      '            'sql = " select ru.U_SDG_ID"
      '            sql = " select ru.U_EXTERNAL_REFERENCE "
      '            sql = sql & " from lims_sys.u_extra_request_user ru,"
      '            sql = sql & "      lims_sys.u_extra_request_data_user rdu"
      '            sql = sql & " where ru.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
      '            sql = sql & " and   rdu.U_EXTRA_REQUEST_DATA_ID=" & dicBarcodeEntities.Keys(i)
      '            Set rs = con.Execute(sql)
      '
      '            strExternalRef = nte(rs("U_EXTERNAL_REFERENCE"))
      '
      '            'get the right SDG
      '            '(considering this might be a revision):
      '            sql = " select d.SDG_ID, d.name"
      '            sql = sql & " from lims_sys.sdg d"
      '            sql = sql & " where d.EXTERNAL_REFERENCE='" & strExternalRef & "'"
      '            sql = sql & " and   instr(d.name, 'V')=0"
      '            sql = sql & " order by d.sdg_id desc"
      '            Set rs = con.Execute(sql)
270               Call sdg_log.InsertLog(CLng(strSdgId), _
                                         "EXTRA.STORAGE", _
                                         CStr(dicBarcodeEntities.Items(i)))
280           End If
290           strLastEntity = dicBarcodeEntities.Items(i)
              
300           Call UpdateArchive(entityType, CStr(dicBarcodeEntities.Items(i)), "F")
310           k = k + 1
320       Next i
330       If k = 1 Then
340           MsgBox "one record was updated"
350       Else
360           MsgBox CStr(k) & " records were updated"
370       End If
          
380       dicBarcodeEntities.RemoveAll
390       cmdReport.Enabled = False
400       Call cmdShowRequests_Click
410       If dicEntities.Count = 0 Then
420           Call lstEntityTypes.RemoveItem(lstEntityTypes.ListIndex)
430       End If
          
440       If lstEntityTypes.ListCount = 0 Then
450           cmdShowRequests.Enabled = False
460       End If
470       lblMicrotom.Visible = False
          
480       Exit Sub
ERR_cmdReport_Click:
490   MsgBox "ERR_cmdReport_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub



Private Function GetSdgForEntity(strEntityName As String, strEntityType As String) As String
500   On Error GoTo ERR_GetSdgForEntity

          Dim rs As Recordset
          Dim sql As String

510       If strEntityType = "Sample" Then
          
520           sql = " select s.SDG_ID"
530           sql = sql & " from lims_sys.sample s"
540           sql = sql & " where s.NAME='" & strEntityName & "'"
              
550       Else
          
560           sql = " select s.SDG_ID"
570           sql = sql & " from lims_sys.sample s,"
580           sql = sql & "      lims_sys.aliquot a"
590           sql = sql & " where a.SAMPLE_ID=s.SAMPLE_ID"
600           sql = sql & " and   a.NAME='" & strEntityName & "'"
          
610       End If

620       Set rs = con.Execute(sql)
          
630       If Not rs.EOF Then
640           GetSdgForEntity = nte(rs("SDG_ID"))
650       End If

660       Exit Function
ERR_GetSdgForEntity:
670   MsgBox "ERR_GetSdgForEntity" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


'clear both dictionaries
'present the selection
'build the dictionary of selected items: ENTITY_NAME -> ROWS_IN_THE_GRID
Private Sub cmdShowRequests_Click()
680   On Error GoTo ERR_cmdShowRequests_Click
          Dim sql As String
          Dim rs As Recordset
          Dim iRows As Integer
          Dim i As Integer
          Dim s As String
          Dim strColorGroup As String
          Dim dicLocations As Dictionary
          
       
690       entityType = lstEntityTypes.Text
700       If entityType = "LBC" Then entityType = "Sample"
          
              
710       Call dicEntities.RemoveAll
720       Call dicBarcodeEntities.RemoveAll
730       Call InitializeGrid
          
          
740           If lstEntityTypes.Text <> "Block" Then
750       sql = "  select rd.U_EXTRA_REQUEST_DATA_ID ID,nvl( du.U_PATHOLAB_NUMBER,'') ||substr(rd.name,11) as PATHOLAB_NAME , rd.NAME ENTITY_NAME,"
760       sql = sql & "         r.NAME ACTION, rdu.U_REQUEST_DETAILS"
770       sql = sql & "         o.NAME, ru.U_CREATED_ON   "
780       sql = sql & "  from lims_sys.u_extra_request_data rd, "
790       sql = sql & "          lims_sys.u_extra_request_data_user rdu, "
800       sql = sql & "          lims_sys.u_extra_request r,"
810       sql = sql & "       lims_sys.u_extra_request_user ru,"
820       sql = sql & "       lims_sys.operator o, "
830       sql = sql & "       lims_sys.sdg_user du "
840       sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
850       sql = sql & "  and r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
860       sql = sql & "  and r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
870       sql = sql & "  and o.OPERATOR_ID=ru.U_CREATED_BY"
880       sql = sql & "  and rdu.U_ENTITY_TYPE='" & lstEntityTypes.Text & "' "
890       sql = sql & "  and rdu.U_STATUS='V' "
900       sql = sql & "  and ru.u_created_on > to_date('01/12/2012', 'dd/mm/yyyy')"
910       sql = sql & "  and du.SDG_ID=ru.U_SDG_ID"
920       sql = sql & "  order by rd.NAME, rd.U_EXTRA_REQUEST_DATA_ID"

930       Set rs = con.Execute(sql)
          
940       Else
          
       'אם הסוג הוא בלוק צריך לבדוק את המיקום שלו (aliquot station)
      'ashi 7/5/13  קריאה 1058
          
950          sql = "select rd.U_EXTRA_REQUEST_DATA_ID ID, nvl(du.U_PATHOLAB_NUMBER,'') ||substr(rd.name,11) as PATHOLAB_NAME , rd.NAME ENTITY_NAME,"
960          sql = sql & "  r.NAME ACTION, rdu.U_REQUEST_DETAILS,"
970          sql = sql & "  o.Name , ru.U_CREATED_ON"
980          sql = sql & "  from lims_sys.u_extra_request_data rd,"
990          sql = sql & "  lims_sys.u_extra_request_data_user rdu,"
1000         sql = sql & "  lims_sys.u_extra_request r,"
1010         sql = sql & "  lims_sys.u_extra_request_user ru,"
1020         sql = sql & "  lims_sys.operator o,"
1030         sql = sql & "  lims_sys.aliquot  al ,"
1040         sql = sql & "  lims_sys.aliquot_user alu, "
1050         sql = sql & "  lims_sys.sdg_user du "
1060         sql = sql & "  Where rd.U_EXTRA_REQUEST_DATA_ID = rdu.U_EXTRA_REQUEST_DATA_ID"
1070         sql = sql & "  and r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
1080         sql = sql & "  and r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
1090         sql = sql & "  and o.OPERATOR_ID=ru.U_CREATED_BY"
1100         sql = sql & "  and rdu.U_ENTITY_TYPE='Block'"
1110         sql = sql & "  and rdu.U_STATUS='V'"
1120         sql = sql & "  and ru.u_created_on > to_date('01/12/2012', 'dd/mm/yyyy')"
1130         sql = sql & "  and du.SDG_ID=ru.U_SDG_ID"
1140         sql = sql & "  AND al.ALIQUOT_ID=alu.ALIQUOT_ID"
1150         sql = sql & "  AND alu.U_ALIQUOT_STATION <'5'"
1160         sql = sql & "  AND al.NAME = SUBSTR(rd.NAME,1,INSTR(rd.NAME,';',1)-1)" 'join to specified aliquot
1170         sql = sql & "  order by rd.NAME, rd.U_EXTRA_REQUEST_DATA_ID"
1180         Set rs = con.Execute(sql)
          'end
1190      End If
1200      iRows = 1
          
1210      While Not rs.EOF
1220          iRows = iRows + 1
          
1230          grid.Rows = iRows
1240          grid.col = 0
1250          grid.row = grid.Rows - 1
              
1260          For i = 0 To rs.Fields.Count - 1
1270              grid.col = i
1280              grid.CellAlignment = vbLeftJustify
                  
1290              s = Trim(CleanSemicolon(nte(rs.Fields(i).Value)))
                 
1300             If nte(rs.Fields(i).Name) = "U_REQUEST_DETAILS" Then
1310                  strColorGroup = GetColorGroup(s)
1320              Else
1330                  strColorGroup = ""
1340              End If
                  
1350              If strColorGroup <> "" Then
1360                  s = s & " (" & strColorGroup & ")"
1370              End If
                  
1380              grid.Text = s
      '            grid.Text = Trim(CleanSemicolon(nte(rs.Fields(i).Value)))
1390          Next i
                         
1400          s = Trim(CleanSemicolon(nte(rs("ENTITY_NAME"))))
                         
1410          If dicEntities.Exists(s) Then
1420              Set dicLocations = dicEntities(s)
1430              Call dicLocations.Add(CStr(grid.row), "")
              
             '     dicEntities(s) = dicEntities(s) & "," & CStr(grid.row)
1440          Else
1450              Set dicLocations = New Dictionary
1460              Call dicLocations.Add(CStr(grid.row), "")
1470              Call dicEntities.Add(s, dicLocations)
      '            Call dicEntities.Add(s, CStr(grid.row))
1480          End If
                         
                         
1490          If ExistRemark(grid.TextMatrix(grid.row, 0)) Then
1500              grid.col = 0
1510              grid.CellFontBold = True
1520          End If
                         
1530          rs.MoveNext
1540      Wend
          
1550      If dicEntities.Count > 0 Then
              'Call InitReportFrame(True)
1560          txtEntityBarcode.Enabled = True
1570          cmdPrint.Enabled = True
1580      End If
          
         ' DEBUG_SHOW_DICTIONARY
          
1590      lblMicrotom.Visible = False
1600      cmdReport.Enabled = False
          
          'for reserve slides - all should be auto-barcoded:
1610      If entityType = "Reserve-Slide" Then
          'And dicOperatorAllowedToReportReserveSlides.Exists(CStr(NtlsUser.GetOperatorId)) Then
1620          Call BarcodeAll
1630      End If
          
1640      Exit Sub
ERR_cmdShowRequests_Click:
1650  MsgBox "ERR_cmdShowRequests_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub





Private Function IExtensionWindow_CloseQuery() As Boolean
          'Happens when the user close the window
          'Call UnloadRequest
1660      IExtensionWindow_CloseQuery = True
End Function

Private Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
1670      IExtensionWindow_DataChange = windowRefreshNone
End Function

Private Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
1680      IExtensionWindow_GetButtons = windowButtonsNone
End Function

Private Sub IExtensionWindow_Internationalise()
End Sub

Private Sub IExtensionWindow_PreDisplay()
1690      Set con = New ADODB.Connection

 Dim cs As String
          cs = NtlsCon.GetADOConnectionString
          
          If NtlsCon.GetServerIsProxy Then
            cs = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
          
    
40        con.Open cs




1710      con.CursorLocation = adUseClient
1720      con.Execute "SET ROLE LIMS_USER"

1730      Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
          
1740      Set sdg_log.con = con
1750      sdg_log.Session = CDbl(NtlsCon.GetSessionId)
End Sub

Private Sub IExtensionWindow_refresh()
    'Code for refreshing the window
End Sub

Private Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)
End Sub

Private Function IExtensionWindow_SaveData() As Boolean
End Function

Private Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)
End Sub

Private Sub IExtensionWindow_SetParameters(ByVal parameters As String)
1760  On Error GoTo ERR_IExtensionWindow_SetParameters

          Dim strMain As String
          Dim s As String
          
1770      strMain = parameters
          
1780      While strMain <> ""
1790          s = getNextStr(strMain, " ")
1800          Call dicOperatorAllowedToReportReserveSlides.Add(s, "")
1810      Wend
          
1820      Exit Sub
ERR_IExtensionWindow_SetParameters:
1830  MsgBox "ERR_IExtensionWindow_SetParameters" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub IExtensionWindow_SetServiceProvider(ByVal serviceProvider As Object)
          Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
1840      Set sp = serviceProvider
1850      Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
1860      Set NtlsCon = sp.QueryServiceProvider("DBConnection")
1870      Set NtlsUser = sp.QueryServiceProvider("User")
End Sub

Private Sub IExtensionWindow_SetSite(ByVal Site As Object)
1880      Set NtlsSite = Site
1890      NtlsSite.SetWindowInternalName ("ExtraRequsts")
1900      NtlsSite.SetWindowRegistryName ("ExtraRequsts")
1910      Call NtlsSite.SetWindowTitle("Pathology Extra Requsts")
End Sub

Private Sub IExtensionWindow_Setup()
1920  On Error GoTo ERR_IExtensionWindow_Setup
          Dim rs As Recordset
          Dim sql As String
      '    Dim sp As SnomedPear
          
          'get snomed T & snomed M pears to search for
          'from the phrase:
      '    Set rs = con.Execute("select phrase_description, phrase_name, phrase_info " & _
      '        " from lims_sys.phrase_entry " & _
      '        " where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
      '        " name = 'BreastCancerReport') " & _
      '        " order by order_number")
      '
      '    While Not rs.EOF
      '        Set sp = New SnomedPear
      '
      '        sp.SnomedT = rs("phrase_description")
      '        sp.SnomedM = rs("phrase_info")
      '
      '        Call dicSnomeds.Add(dicSnomeds.Count, sp)
      '
      '        rs.MoveNext
      '    Wend
      '
      '    Call InitializeGrid
      '    Call InitDates
      '    Call MaskEdBoxFrom.SetFocus
          

      '    Set rs = con.Execute(" select distinct rdu.U_ENTITY_TYPE " & _
      '                         " from lims_sys.u_extra_request_data_user rdu " & _
      '                         " order by rdu.U_ENTITY_TYPE ")
      '    While Not rs.EOF
      '        lstEntityTypes.AddItem (rs("U_ENTITY_TYPE"))
      '
      '        rs.MoveNext
      '    Wend
      '    lstEntityTypes.Selected(0) = True
          
      '    Call InitReportFrame(False)
1930      cmdReport.Enabled = False
1940      cmdPrint.Enabled = False
1950      txtEntityBarcode.Enabled = False
1960      Call InitializeGrid
1970      Call InitEntityTypes
          
          
1980      Exit Sub
ERR_IExtensionWindow_Setup:
1990  MsgBox "ERR_IExtensionWindow_Setup" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
2000      IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub ConnectSameSession(ByVal aSessionID)
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
2010      aProc.ActiveConnection = con
2020      aProc.CommandText = "lims.lims_env.connect_same_session"
2030      aProc.CommandType = adCmdStoredProc

2040      aSession.Type = adDouble
2050      aSession.Direction = adParamInput
2060      aSession.Value = aSessionID
2070      aProc.parameters.Append aSession

2080      aProc.Execute
2090      Set aSession = Nothing
2100      Set aProc = Nothing
End Sub

Private Sub IExtensionWindow2_Close()
End Sub


Private Sub InitializeGrid()
          Dim x As Integer
          Dim s As String
              
2110      grid.Clear
              
2120      grid.AllowBigSelection = False
2130      grid.AllowUserResizing = flexResizeNone
2140      grid.Enabled = True
          
2150      grid.ScrollBars = flexScrollBarBoth
2160      grid.SelectionMode = flexSelectionFree
2170      grid.AllowUserResizing = flexResizeBoth

2180      grid.Font.Size = 12
2190      grid.Rows = 2
2200      grid.Cols = 7
2210      grid.FixedRows = 1
2220      grid.FixedCols = 0

2230      grid.row = 0
2240      grid.RowHeight(x) = 400
      '    For X = 1 To grid.Rows - 1
      '        grid.row = X
      '        grid.RowHeight(X) = 600
      '    Next X
2250          grid.col = 0
2260          grid.ColWidth(1) = 1400
2270      For x = 1 To 3
2280          grid.col = x
2290          grid.ColWidth(x) = 2200
2300      Next x

2310      For x = 4 To grid.Cols - 1
2320          grid.col = x
2330          grid.ColWidth(x) = 2400
2340      Next x
          
          'set the text for the COLUMN HEADERS:
2350      grid.row = 0
2360      grid.col = 0
      '    grid.CellAlignment = vbLeftJustify
      '    grid.Text = "Entity Type"

2370      grid.CellAlignment = vbLeftJustify
2380      grid.Text = "ID"
          
2390      grid.col = grid.col + 1
2400      grid.CellAlignment = vbLeftJustify
2410      grid.Text = "Patho-Lab Name"

2420      grid.col = grid.col + 1
2430      grid.CellAlignment = vbLeftJustify
2440      grid.Text = "Entity Name"

2450      grid.col = grid.col + 1
2460      grid.CellAlignment = vbLeftJustify
2470      grid.Text = "Action"
          
2480      grid.col = grid.col + 1
2490      grid.CellAlignment = vbLeftJustify
2500      grid.Text = "Details"
          
2510      grid.col = grid.col + 1
2520      grid.CellAlignment = vbLeftJustify
2530      grid.Text = "Created By"
          
2540      grid.col = grid.col + 1
2550      grid.CellAlignment = vbLeftJustify
2560      grid.Text = "Created On"

End Sub

'used for a concatenated fieled: value1;value2
'gets the 1st value of that string
Private Function CleanSemicolon(str As String) As String
          Dim i As Integer
          
2570      CleanSemicolon = str
          
2580      i = InStr(1, str, ";")
2590      If i = 0 Then Exit Function
          
2600      CleanSemicolon = Left(str, i - 1)
End Function


Private Function nte(e As Variant) As Variant
2610      nte = IIf(IsNull(e), "", e)
End Function


Private Sub lstEntityTypes_Click()
2620      cmdShowRequests.Enabled = True
2630      cmdShowRequests_Click
End Sub

'if a new request is made, of a type we do not currently
'have in the list of types (slide / block / sample etc.)
'it is added to the list:
Private Sub Timer1_Timer()
2640      Timer1.Enabled = False
2650      Call RefreshEntityTypes
2660      Timer1.Enabled = True
End Sub

'mark the line(s) of this entity
'get the EXTRA_REQUSER_DATA_ID at all the entries this entity is found
'and add it to the dictionaty of barcoded entities
Private Sub txtEntityBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
2670  On Error GoTo ERR_txtEntityBarcode_KeyDown
          Dim strEntityName As String

2680      If KeyCode <> vbKeyReturn Then Exit Sub

2690      strEntityName = UCase(txtEntityBarcode.Text)

2700      Call BarcodeEntity(strEntityName)

          'show the microtom for blocks:
      '    If lstEntityTypes.Text = "Block" Then
      '        lblMicrotom.Caption = "Microtome: " & GetMicrotom(strEntityName)
      '        lblMicrotom.Visible = True
      '    Else
      '        lblMicrotom.Visible = False
      '    End If
          
2710      txtEntityBarcode.Text = ""
          

2720      Exit Sub
ERR_txtEntityBarcode_KeyDown:
2730  MsgBox "ERR_txtEntityBarcode_KeyDown" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub BarcodeEntity(strEntityName As String)
2740  On Error GoTo ERR_BarcodeEntity
          Dim i As Integer
          Dim k As Integer
          Dim s As String
          Dim MainStr As String
          Dim dicLocations As Dictionary



2750      If Not dicEntities.Exists(strEntityName) Then
2760          MsgBox "The barcode value doesn't exist in the list"
2770          txtEntityBarcode.Text = ""
2780          txtEntityBarcode.SetFocus
2790          Exit Sub
2800      End If

2810      Set dicLocations = dicEntities(strEntityName)
      '    MainStr = dicEntities(strEntityName)
          
2820      lblMicrotom.Visible = False
          
2830      For i = 0 To dicLocations.Count - 1
2840          s = dicLocations.Keys(i)
2850          grid.row = CLng(s)
2860          grid.col = 0

2870          Call ShowMicrotom(entityType, grid.TextMatrix(grid.row, 2), _
                                                     grid.TextMatrix(grid.row, 0), _
                                                     strEntityName)

2880          If Not dicBarcodeEntities.Exists(grid.Text) Then
2890              Call dicBarcodeEntities.Add(grid.Text, strEntityName)
2900          End If

2910          For k = 0 To grid.Cols - 1
2920              grid.col = k
2930              grid.CellBackColor = MARK_SELECTED
2940          Next k
2950      Next i
          
      '    While MainStr <> ""
      '        s = getNextStr(MainStr, ",")
      '
      '        grid.row = CLng(s)
      '        grid.col = 0
      '
      '        Call ShowMicrotom(lstEntityTypes.Text, grid.TextMatrix(grid.row, 2), _
      '                                               grid.TextMatrix(grid.row, 0), _
      '                                               strEntityName)
      '
      '        If Not dicBarcodeEntities.Exists(grid.Text) Then
      '            Call dicBarcodeEntities.Add(grid.Text, strEntityName)
      '        End If
      '
      '        For i = 0 To grid.Cols - 1
      '            grid.col = i
      '            grid.CellBackColor = MARK_SELECTED
      '        Next i
      '    Wend
          
2960      cmdReport.Enabled = True

2970      Exit Sub
ERR_BarcodeEntity:
2980  MsgBox "ERR_BarcodeEntity" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'show the microtom for adding slide to the block if
'for the sub entity the DESCRIPTION is T
'(the doctor asked for it to be cut by the same microtom as the last one)
Private Sub ShowMicrotom(strEntityType As String, strAction As String, _
                         strExtraRequestDataId As String, _
                         strEntityName As String)
2990  On Error GoTo ERR_ShowMicrotom
          Dim rs As Recordset
          Dim sql As String

3000      If strEntityType <> "Block" Then Exit Sub
          
3010      If strAction <> "Add Slide" Then Exit Sub
          
3020      sql = " select rd.DESCRIPTION"
3030      sql = sql & " from lims_sys.u_extra_request_data rd"
3040      sql = sql & " where rd.U_EXTRA_REQUEST_DATA_ID='" & strExtraRequestDataId & "'"
              
3050      Set rs = con.Execute(sql)
          
3060      If rs.EOF Then Exit Sub

3070      If nte(rs("DESCRIPTION")) = "T" Then
3080          lblMicrotom.Caption = "Microtome: " & GetMicrotom(strEntityName)
3090          lblMicrotom.Visible = True
3100      End If
              
3110      Exit Sub
ERR_ShowMicrotom:
3120  MsgBox "ERR_ShowMicrotom" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function getNextStr(ByRef s As String, c As String) As String
          Dim p
          Dim res
3130      p = InStr(1, s, c)
3140      If (p = 0) Then
3150          res = s
3160          s = ""
3170          getNextStr = res
3180      Else
3190          res = Mid$(s, 1, p - 1)
3200          s = Mid$(s, p + Len(c), Len(s))
3210          getNextStr = res
3220      End If
End Function

'debug-print the collection of all entities at the current selection
Private Sub DEBUG_SHOW_DICTIONARY()
3230  On Error GoTo ERR_DEBUG_SHOW_DICTIONARY
          Dim i As Integer
          Dim j As Integer
          Dim s As String
          Dim MainStr As String
          

3240      For i = 0 To dicEntities.Count - 1
3250          s = "_" & dicEntities.Keys(i) & "_"
              
3260          MainStr = dicEntities.Items(i)
              
3270          While MainStr <> ""
3280              s = s & vbCrLf & getNextStr(MainStr, ",")
3290          Wend
              
3300          MsgBox s
3310      Next i

3320      Exit Sub
ERR_DEBUG_SHOW_DICTIONARY:
3330  MsgBox "ERR_DEBUG_SHOW_DICTIONARY" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub PrintToPrinter(iStartRow As Integer, iEndRow As Integer, iRows As Integer)
3340  On Error GoTo ERR_PrintToPrinter

          Dim BottomOffset As Long
          Dim RightOffset As Long
          Dim strName As String
          Dim s As String
          


          Dim i As Integer, j As Integer   ' Declare variables
          
          
          
          
3350      Printer.ScaleMode = 3   ' Set ScaleMode to pixels.
3360      Printer.Orientation = vbPRORLandscape
          
3370      TopOffset = 400 'Int(Printer.ScaleHeight / 8)
3380      BottomOffset = 50 'Int(Printer.ScaleHeight / 12)
3390      LeftOffset = 100 'CInt(Printer.ScaleHeight / 12)
3400      RightOffset = 50
3410      ColSize = Int((Printer.ScaleWidth - LeftOffset - RightOffset) / (grid.Cols))
3420      RowSize = Int((Printer.ScaleHeight - TopOffset - BottomOffset) / (iRows))
3430      Printer.DrawWidth = 2   ' Set DrawWidth.
          
3440      Printer.Font = "Arial"
3450      Printer.FontSize = 10
3460      Printer.FontBold = False
3470      Printer.FontUnderline = False
3480      Printer.CurrentY = Int(TopOffset / 2)
3490      Printer.CurrentX = LeftOffset
          'to be modified:
          'Printer.Print "Plate Name: " & strPlateName '"Plate Name"
3500      Printer.FontUnderline = False


          'print report name and dates above the table:
3510      Printer.CurrentY = Int(TopOffset / 6)
3520      Printer.CurrentX = Int(LeftOffset)
3530      Printer.Print "Entity Type: " & lstEntityTypes.Text

3540      Printer.FontSize = 8

          'print columns headers:
3550      grid.row = 0
3560      For i = 0 To grid.Cols - 1
3570          grid.col = i
3580          strName = Left(grid.Text, MAX_DIGITS_PER_CELL)
              
3590          Printer.CurrentY = Int(TopOffset / 1.5)
              'Printer.CurrentY = 3 * Int(TopOffset / 4)
              
3600          Printer.CurrentX = ColPixel(i)
              'Printer.CurrentX = ColPixel(i) + Int(ColSize / 2) - 50
              
3610          Printer.Print strName
      '        While strName <> ""
      '            s = getNextStr(strName, vbCrLf)
      '            Printer.Print Left(s, 9)
      '            Printer.CurrentX = ColPixel(i)
      '            Printer.CurrentY = Printer.CurrentY + 10
      '
      '            If Len(s) > 9 Then
      '                Printer.Print Mid(s, 10, 9)
      '                Printer.CurrentX = ColPixel(i)
      '                Printer.CurrentY = Printer.CurrentY + 10
      '            End If
      '        Wend
3620      Next i
          
          'print row headers:
      '    grid.col = 0
      '    For i = 0 To grid.Rows - 1
      '        grid.row = i
      '        strName = grid.Text
      '        Printer.CurrentX = Int(LeftOffset / 4)
      '        Printer.CurrentY = RowPixel(i) - 20
      '
      '        Printer.Print strName
      '
      ''        strName = Mid(strName, InStr(1, strName, " ", vbTextCompare) + 1)
      ''
      ''        While strName <> ""
      ''            s = getNextStr(strName, " ")
      ''            Printer.Print Left(s, 15)
      ''            Printer.CurrentX = Int(LeftOffset / 4)
      ''            Printer.CurrentY = Printer.CurrentY + 10
      ''        Wend
      '    Next i

          'try a diff font - fixed sized one
          'to do -
          '1. count how many digits can enter in a cell
          '2. change font size if needed:
          
3630      strFontName = "Miriam Fixed"
3640      iFontSize = 8
3650      isFontBold = False
          
3660      Printer.Font = strFontName
3670      Printer.FontSize = iFontSize
3680      Printer.FontBold = isFontBold

3690      i = TopOffset
3700      While i <= Printer.ScaleHeight - BottomOffset
3710          Printer.Line (LeftOffset, i)-(Printer.ScaleWidth - RightOffset, i)
3720          i = i + RowSize
3730      Wend
3740      i = LeftOffset
          

3750      While i <= Int(Printer.ScaleWidth) + 1 - RightOffset
3760          Printer.Line (i, TopOffset)-(i, Printer.ScaleHeight - BottomOffset)
3770          i = i + ColSize
3780      Wend

3790      For i = 0 To grid.Cols - 1
3800          grid.col = i
          
3810          For j = iStartRow To iEndRow
                  Dim iNameSize As Integer
                  Dim iNameSizeInPixels As Integer
                  Dim iCenteredColPixel As Integer
                  Dim iCellLeftShift As Integer
                  
                  'change to udi report
                  '15.05.2006
                  'get the text from that cell in the grid
                  'strName=...
                  
3820              grid.row = j
                  
3830              strName = Left(grid.Text, MAX_DIGITS_PER_CELL)
      '            strName = Replace(strName, vbCrLf, " ")
      '            strName = aAliquotArray(j, i).strName
3840              iNameSize = Len(strName)
                  
      '            iNameSizeInPixels = iNameSize * ColSize / iMaxDigitsPerCell
                     
                  'in case the name is too long
                  'there in no shift at all:
      '            If iNameSizeInPixels > ColSize Then
      '                iNameSizeInPixels = ColSize
      '            End If
                  
                  'number of blank pixels from the left of the cell
                  'for this name:
3850              iCellLeftShift = 0 ' (ColSize - iNameSizeInPixels) / 2
                  
                  'start printing this cell's name
                  'at that pixel:
            '      iCenteredColPixel = iCellLeftShift + ColPixel(i) '+ iNameSizeInPixels / 4
                  
            '      Printer.CurrentX = iCenteredColPixel
                  
3860              Printer.CurrentX = ColPixel(i)
3870              Printer.CurrentY = RowPixel(j - iStartRow) - 20
                  
3880              Printer.Print strName
      '            While strName <> ""
      '                s = getNextStr(strName, vbCrLf)
      '                Printer.Print Left(s, 8)
      '                Printer.CurrentX = ColPixel(i)
      '                Printer.CurrentY = Printer.CurrentY + 10
      '            Wend
                  
                  
                  'not all the name is printed if there is not
                  'enough space:
                  
                  
                  'Printer.Print Left(strName, iMaxDigitsPerCell)
                  
                  
                  'Printer.Print Left(aAliquotArray(j, i), 11)
3890          Next j
3900      Next i
          'Printer.KillDoc
3910      Printer.EndDoc
          
3920      Exit Sub
ERR_PrintToPrinter:
3930  MsgBox "ERR_PrintToPrinter" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub



Private Function ColPixel(col As Integer) As Long
3940      ColPixel = LeftOffset + col * ColSize + 10
End Function


Private Function RowPixel(row As Integer) As Long
3950      RowPixel = TopOffset + row * RowSize + Int(RowSize / 2)
          'RowPixel = TopOffset + row * RowSize + Int(RowSize / 2) - 50
End Function

'get list of entity types in status V
'(not yet reported as out of archive)
Private Sub InitEntityTypes()
3960  On Error GoTo ERR_InitEntityTypes
          Dim rs As Recordset

3970      Set rs = con.Execute(" select distinct rdu.U_ENTITY_TYPE " & _
                               " from lims_sys.u_extra_request_data_user rdu " & _
                               " where rdu.u_status='V' order by rdu.U_ENTITY_TYPE ")
3980      While Not rs.EOF
3990          lstEntityTypes.AddItem (rs("U_ENTITY_TYPE"))

4000          rs.MoveNext
4010      Wend
          
4020      If lstEntityTypes.ListCount > 0 Then
4030          lstEntityTypes.Selected(0) = True
4040          cmdShowRequests.Enabled = True
4050          Call cmdShowRequests_Click
4060      Else
4070          cmdShowRequests.Enabled = False
4080      End If

4090      Exit Sub
ERR_InitEntityTypes:
4100  MsgBox "ERR_InitEntityTypes" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'add to the list of entity types:
Private Sub RefreshEntityTypes()
4110  On Error GoTo ERR_RefreshEntityTypes
          Dim rs As Recordset
          Dim strExistingTypes As String
          Dim i As Integer
          
          'in order for the sql statement to be valid
          'where there are no items in the list of Entity Types:
4120      strExistingTypes = ",'xyzw'"
          
          
          'get existing entity types:
4130      For i = 0 To lstEntityTypes.ListCount - 1
4140          strExistingTypes = ",'" & lstEntityTypes.List(i) & "'" & strExistingTypes
4150      Next i
4160      strExistingTypes = Mid(strExistingTypes, 2)

4170      Set rs = con.Execute(" select distinct rdu.U_ENTITY_TYPE " & _
                               " from lims_sys.u_extra_request_data_user rdu " & _
                               " where rdu.u_status='V' " & _
                               " and rdu.U_ENTITY_TYPE NOT IN (" & strExistingTypes & ") " & _
                               " order by rdu.U_ENTITY_TYPE ")
4180      While Not rs.EOF
4190          lstEntityTypes.AddItem (rs("U_ENTITY_TYPE"))

4200          rs.MoveNext
4210      Wend
          
      '    If lstEntityTypes.ListCount = 1 Then
      '        cmdShowRequests.Enabled = True
      '     End If

4220      Exit Sub
ERR_RefreshEntityTypes:
4230  MsgBox "ERR_RefreshEntityTypes" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'get the letter representing the last workstation to use
'SlideGenerationReport for this block:
Private Function GetMicrotom(strBlockName As String)
4240  On Error GoTo ERR_GetMicrotom
          Dim sql As String
          Dim rs As Recordset
          Dim strWorkstationId As String
          Dim strWorkstationName As String

          'get the workstation id:
4250      sql = " select au.U_LAST_MICROTOME"
4260      sql = sql & " from lims_sys.aliquot a,lims_sys.aliquot_user au"
4270      sql = sql & " where a.ALIQUOT_ID=au.ALIQUOT_ID"
4280      sql = sql & " and a.NAME='" & strBlockName & "'"

4290      Set rs = con.Execute(sql)
4300      strWorkstationId = nte(rs("U_LAST_MICROTOME"))
4310      If strWorkstationId = "" Then Exit Function
          
          'get the workstation name:
4320      sql = " select name "
4330      sql = sql & " from lims_sys.workstation "
4340      sql = sql & " where workstation_id = " & strWorkstationId
          
4350      Set rs = con.Execute(sql)
4360      strWorkstationName = rs("name")

          'get the microtom for the workstation name:
4370      Set rs = con.Execute("select phrase_description from lims_sys.phrase_entry " & _
              "where phrase_name = '" & strWorkstationName & "' and " & _
              "phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'StationToMicrotom')")
              
4380      If Not rs.EOF Then
4390          GetMicrotom = nte(rs("phrase_description"))
4400      End If
          
4410      Exit Function
ERR_GetMicrotom:
4420  MsgBox "ERR_GetMicrotom" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'a click on a cell shows the remark for the request
'containing this entity
Private Sub grid_DblClick()
4430  On Error GoTo ERR_grid_Click
          Dim strRequestDataId As String
          
4440      strRequestDataId = grid.TextMatrix(grid.row, 0)

4450      Call frmRemarks.Initialize(con, strRequestDataId)
4460      Call frmRemarks.Show(vbModal)
          
4470      Exit Sub
ERR_grid_Click:
4480  MsgBox "MsgBox" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description

End Sub

Private Function ExistRemark(strExtraRequestDataId As String) As Boolean
4490  On Error GoTo ERR_ExistRemark
          Dim rs As Recordset
          Dim sql As String
          
4500      sql = " select r.DESCRIPTION "
4510      sql = sql & " from lims_sys.u_extra_request_data rd, "
4520      sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
4530      sql = sql & "      lims_sys.u_extra_request r"
4540      sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
4550      sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
4560      sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strExtraRequestDataId

4570      Set rs = con.Execute(sql)
          
4580      If nte(rs("DESCRIPTION")) = "" Then
4590          ExistRemark = False
4600      Else
4610          ExistRemark = True
4620      End If

4630      Exit Function
ERR_ExistRemark:
4640  MsgBox "ERR_ExistRemark" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


'get the shortcut for the color group of this color
'an empty string is returned if this is not on of the known colors
Private Function GetColorGroup(strColor As String) As String
4650  On Error GoTo ERR_GetColorGroup
          Dim rs As Recordset
          Dim sql As String
          Dim s As String
          
4660      sql = " select ph.NAME "
4670      sql = sql & " from lims_sys.phrase_header ph,"
4680      sql = sql & "      lims_sys.phrase_entry pe"
4690      sql = sql & " where pe.PHRASE_ID=ph.PHRASE_ID"
4700      sql = sql & " and pe.PHRASE_NAME='" & strColor & "' "
4710      sql = sql & " and ph.NAME in ('Pathology Molecular Stains',"
4720      sql = sql & "                 'Pathology Special Stains',"
4730      sql = sql & "                 'Pathology Other Stain Options',"
4740      sql = sql & "                 'Pathology Imonohistochemistry stains')"

4750      Set rs = con.Execute(sql)
          
4760      If rs.EOF = True Then Exit Function
          
4770      s = nte(rs("NAME"))
          
4780      Select Case s
              Case "Pathology Molecular Stains"
4790              s = "Mol"
4800          Case "Pathology Special Stains"
4810              s = "S"
4820          Case "Pathology Imonohistochemistry stains"
4830              s = "IHC"
4840          Case "Pathology Other Stain Options"
4850              s = "O"
4860      End Select

4870      GetColorGroup = s

4880      Exit Function
ERR_GetColorGroup:
4890  MsgBox "ERR_GetColorGroup" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'indicate if this item (sample / block / slide)
'is in the tissue archive or not:
Private Sub UpdateArchive(strEntityType As String, strEntityName As String, strStored As String)
4900  On Error GoTo ERR_UpdateArchive
          Dim sql As String
              
4910      Select Case strEntityType
              Case "Sample"
4920              sql = " update lims_sys.sample_user su"
4930              sql = sql & " set su.U_ARCHIVE='" & strStored & "'"
4940              sql = sql & " where su.SAMPLE_ID="
4950              sql = sql & " ("
4960              sql = sql & "   select s.SAMPLE_ID"
4970              sql = sql & "   from lims_sys.sample s"
4980              sql = sql & "   where s.NAME='" & strEntityName & "'"
4990              sql = sql & " )"
              
5000          Case "Block"
5010              sql = " update lims_sys.aliquot_user au"
5020              sql = sql & " set au.U_ARCHIVE='" & strStored & "'"
5030              sql = sql & " where au.ALIQUOT_ID="
5040              sql = sql & " ("
5050              sql = sql & "   select a.ALIQUOT_ID"
5060              sql = sql & "   from lims_sys.aliquot a"
5070              sql = sql & "   where a.NAME='" & strEntityName & "'"
5080              sql = sql & " )"
5090              sql = sql & " and exists"
5100              sql = sql & " ("
5110              sql = sql & "   select 1 "
5120              sql = sql & "   from lims_sys.aliquot_formulation af"
5130              sql = sql & "   where af.PARENT_ALIQUOT_ID=au.ALIQUOT_ID"
5140              sql = sql & " )"
              
5150          Case Else
5160              sql = " update lims_sys.aliquot_user au"
5170              sql = sql & " set au.U_ARCHIVE='" & strStored & "'"
5180              sql = sql & " where au.ALIQUOT_ID="
5190              sql = sql & " ("
5200              sql = sql & "   select a.ALIQUOT_ID"
5210              sql = sql & "   from lims_sys.aliquot a"
5220              sql = sql & "   where a.NAME='" & strEntityName & "'"
5230              sql = sql & " )"
5240              sql = sql & " and exists"
5250              sql = sql & " ("
5260              sql = sql & "   select 1 "
5270              sql = sql & "   from lims_sys.aliquot_formulation af"
5280              sql = sql & "   where af.CHILD_ALIQUOT_ID=au.ALIQUOT_ID"
5290              sql = sql & " )"
                  
5300      End Select
          
          
      '    If strEntityType = "Sample" Then
      '        sql = " update lims_sys.sample_user su"
      '        sql = sql & " set su.U_ARCHIVE='" & strStored & "'"
      '        sql = sql & " where su.SAMPLE_ID="
      '        sql = sql & " ("
      '        sql = sql & "   select s.SAMPLE_ID"
      '        sql = sql & "   from lims_sys.sample s"
      '        sql = sql & "   where s.NAME='" & strEntityName & "'"
      '        sql = sql & " )"
      '    Else
      '        sql = " update lims_sys.aliquot_user au"
      '        sql = sql & " set au.U_ARCHIVE='" & strStored & "'"
      '        sql = sql & " where au.ALIQUOT_ID="
      '        sql = sql & " ("
      '        sql = sql & "   select a.ALIQUOT_ID"
      '        sql = sql & "   from lims_sys.aliquot a"
      '        sql = sql & "   where a.NAME='" & strEntityName & "'"
      '        sql = sql & " )"
      '    End If
          
5310      Call con.Execute(sql)

5320      Exit Sub
ERR_UpdateArchive:
5330  MsgBox "UpdateArchive" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

