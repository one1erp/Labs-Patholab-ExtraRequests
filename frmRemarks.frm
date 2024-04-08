VERSION 5.00
Begin VB.Form frmRemarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "הערות"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemark 
      Alignment       =   1  'Right Justify
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private con As Connection



Public Sub Initialize(con_ As Connection, strRequestDataId As String)
5340  On Error GoTo ERR_Initialize
          Dim rs As Recordset
          Dim sql As String
          
5350      Set con = con_
          
5360      sql = " select r.DESCRIPTION "
5370      sql = sql & " from lims_sys.u_extra_request_data rd, "
5380      sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
5390      sql = sql & "      lims_sys.u_extra_request r"
5400      sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
5410      sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
5420      sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strRequestDataId

5430      Set rs = con.Execute(sql)
          
5440      If rs.EOF = True Then Exit Sub

5450      txtRemark.Text = nte(rs("DESCRIPTION"))

5460      Exit Sub
ERR_Initialize:
5470  MsgBox "ERR_Initialize" & vbCrLf & Err.Description
End Sub



Private Function nte(e As Variant) As Variant
5480      nte = IIf(IsNull(e), "", e)
End Function

Private Sub Form_Click()
5490      Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
5500      If KeyAscii = vbKeyEscape Then
5510          Me.Hide
5520      End If
End Sub

Private Sub txtRemark_Click()
5530      Me.Hide
End Sub
