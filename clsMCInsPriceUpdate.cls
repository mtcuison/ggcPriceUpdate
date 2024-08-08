VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCInsPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Update MC Installment Price
'
' Copyright 2007 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 05/22/2009 02:19 pm ]
'     Start coding this object...
'  XerSys [ 06/22/2009 10:23 am ]
'     Continue coding this object.
'  XerSys [ 09/29/2011 02:46 am ]
'     Added Ending of Mortgage Fee.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMCInsPriceUpdate"
Private Const pxeChangeFrom As String = "1"

Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset
Private p_oActRte As Recordset
Private p_oHistry As Recordset

Private p_sSQLModel As String
Private p_sSQLActRt As String

Private p_sMCCatIDx As String
Private p_sMCCatNme As String
Private p_dInsPrice As Date
Private p_sApproved As String

Private p_nMiscChrg As Double
Private p_nRebatesx As Double
Private p_nEndMrtgg As Double

Private p_sBranchCd As String
Private p_sParentxx As String

Private pbInitTran As Boolean
Private pbMasModified As Boolean
Private pbDetModified As Boolean
Private pbRteModified As Boolean

Private pbShowMsg As Boolean

Public Event DetailRetrieved()
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get TermRate(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Row > (p_oActRte.RecordCount - 1) Then Exit Property
   p_oActRte.Move Row, adBookmarkFirst
   
   TermRate = p_oActRte(Index)
End Property

Property Let TermRate(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   With p_oActRte
      If Row > .RecordCount Then
         Exit Property
      ElseIf Row = .RecordCount Then
         ' check first if last record have a valid parts id
         If CInt(.Fields("nAcctTerm")) = 0 Then
            MsgBox "Inserting Empty Record is Not Allowed!" & vbCrLf & vbCrLf & _
                  "Please inform the SEG/SSG of Guanzon Group of Companies!", vbCritical, "Warning"
            Exit Property
         End If
         
         Call AddTermRate
      Else
         .Move Row, adBookmarkFirst
      End If
   
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      Case 1, 2, "nacctterm", "nacctthru"
         If Not IsNumeric(Value) Then Exit Property
         
         If pbRteModified = False Then
            pbRteModified = .Fields(Index) = CInt(Value)
         End If
         
         .Fields(Index) = CInt(Value)
      Case 3, "nfactorrt"
         If Not IsNumeric(Value) Then Exit Property
         
         .Fields(Index) = CDbl(Value)
         Debug.Print Value
      End Select
   End With
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   p_oDetail.Move Row, adBookmarkFirst
   
   Detail = p_oDetail(Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   With p_oDetail
      If Row > .RecordCount Then
         Exit Property
      ElseIf Row = .RecordCount Then
         ' check first if last record have a valid parts id
         If .Fields("sModelIDx") = Empty Then
            MsgBox "Inserting Empty Record is Not Allowed!" & vbCrLf & vbCrLf & _
                  "Please inform the SEG/SSG of Guanzon Group of Companies!", vbCritical, "Warning"
            Exit Property
         End If
         
         Call AddDetail
      Else
         .Move Row, adBookmarkFirst
      End If
   
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      Case 1, "smodelidx"
         Call getDetail(Value, False)
      Case 2, "smodelcde"
      Case 3, "sbrandnme"
      Case 4, 5, "nselprice", "nmindownx"
         If Not IsNumeric(Value) Then Exit Property
         
         If pbDetModified = False Then
            pbDetModified = p_oDetail(Index) = CDbl(Value)
         End If
         
         p_oDetail(Index) = CDbl(Value)
      End Select
   End With
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "smccatidx"
      Master = p_sMCCatIDx
   Case 1, "smccatnme"
      Master = p_sMCCatNme
   Case 2, "dpricexxx"
      Master = p_dInsPrice
   Case 3, "nmiscchrg"
      Master = p_nMiscChrg
   Case 4, "nrebatesx"
      Master = p_nRebatesx
   Case 5, "nendmrtgg"
      Master = p_nEndMrtgg
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "smccatidx"
   Case 1, "smccatnme"
      Call getMaster(Value, False)
   Case 2, "dpricexxx"
      If Not IsDate(Value) Then Exit Property
   
      p_dInsPrice = Value
   Case 3, "nmiscchrg"
      If Not IsNumeric(Value) Then Exit Property
      
      If pbMasModified = False Then
         pbMasModified = p_nMiscChrg <> CDbl(Value)
      End If
      Debug.Print pbMasModified
      
      p_nMiscChrg = CDbl(Value)
   Case 4, "nrebatesx"
      If Not IsNumeric(Value) Then Exit Property
      
      If pbMasModified = False Then
         pbMasModified = p_nRebatesx <> CDbl(Value)
      End If
      
      Debug.Print pbMasModified
      p_nRebatesx = CDbl(Value)
   Case 5, "nendmrtgg"
      If Not IsNumeric(Value) Then Exit Property
      
      If pbMasModified = False Then
         pbMasModified = p_nEndMrtgg <> CDbl(Value)
      End If
      
      Debug.Print pbMasModified
      p_nEndMrtgg = CDbl(Value)
   End Select
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Property Get ItemCount() As Integer
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oDetail.RecordCount
End Property

Property Get RateCount() As Integer
   If pbInitTran = False Then Exit Property
   
   RateCount = p_oActRte.RecordCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQLBrows As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sSQLModel = "SELECT" & _
                     "  a.sModelIDx" & _
                     ", b.sModelNme" & _
                     ", b.sModelCde" & _
                     ", d.sBrandNme" & _
                     ", a.nSelPrice" & _
                     ", a.nMinDownx" & _
                     ", c.sMCCatNme" & _
                     ", a.sApproved" & _
                     ", a.dInsPrice" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                     ", a.sMCCatIDx" & _
                     ", a.cRecdStat" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                     ", MC_Category c" & _
                     ", Brand d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sMCCatIDx = c.sMCCatIDx" & _
                     " AND b.sBrandIDx = d.sBrandIDx" & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                  " ORDER BY b.sModelNme"
   
   p_sSQLActRt = "SELECT" & _
                     "  sMCCatIDx" & _
                     ", nAcctTerm" & _
                     ", nAcctThru" & _
                     ", nFactorRt" & _
                     ", sApproved" & _
                     ", dPricexxx" & _
                     ", sModified" & _
                     ", dModified" & _
                  " FROM MC_Term_Category" & _
                  " ORDER BY nAcctTerm"

   p_sMCCatIDx = ""
   p_sMCCatNme = ""
   p_nMiscChrg = 0#
   p_nRebatesx = 0#
   p_nEndMrtgg = 0#
   p_dInsPrice = p_oAppDrivr.SysDate
   p_sApproved = ""
   
   Set p_oDetail = New Recordset
   p_oDetail.Open AddCondition(p_sSQLModel, "0 = 1"), p_oAppDrivr.Connection, , , adCmdText
   
   Set p_oActRte = New Recordset
   p_oActRte.Open AddCondition(p_sSQLActRt, "0 = 1"), p_oAppDrivr.Connection, , , adCmdText
   
   pbInitTran = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "p_oBaseClas_WillSave"
   On Error GoTo errProc
   
   If MsgBox("Are you sure you want to Save this Transaction?", _
         vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
   
   lsSQL = AddCondition(p_sSQLModel, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
   Set p_oHistry = New Recordset
   p_oHistry.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   p_oHistry.Sort = "sModelIDx"
   p_oDetail.Sort = "sModelIDx"
   p_oDetail.MoveFirst
   
   p_oAppDrivr.BeginTrans
reCompare:
   Do While p_oDetail.EOF = False
      If p_oHistry.EOF = False Then
         Debug.Print p_oDetail("sModelIDx"), p_oHistry("sModelIDx")
         Select Case StrComp(p_oDetail("sModelIDx"), p_oHistry("sModelIDx"), vbTextCompare)
         Case -1 ' Modified < Original
            If addNewModel() = False Then GoTo endWithRoll
         Case 0 ' Equal
            If IsNull(p_oHistry("nMinDownx")) Then
               If updateModelPrice() = False Then GoTo endWithRoll
            ElseIf p_oDetail("nSelPrice") <> p_oHistry("nSelPrice") Or _
                  p_oDetail("nMinDownx") <> p_oHistry("nMinDownx") Then
               If updateModelPrice() = False Then GoTo endWithRoll
            End If
            p_oHistry.MoveNext
         Case 1  ' Modified > original
            With p_oHistry
               Do
                  If StrComp(p_oDetail.Fields("sModelIDx"), _
                        p_oHistry("sModelIDx"), vbTextCompare) = 1 Then
                     If removeModel() = False Then GoTo endWithRoll
                  Else
                     Exit Do
                  End If
                  
                  .MoveNext
               Loop Until .EOF
            End With
            GoTo reCompare
         End Select
      Else
         If addNewModel() = False Then
            GoTo endWithRoll
         End If
      End If
      
      p_oDetail.MoveNext
   Loop
      
   With p_oHistry
      Do Until .EOF
         If removeModel() = False Then GoTo endWithRoll
         
         .MoveNext
      Loop
   End With
   
   ' now check term/rate record
   lsSQL = AddCondition(p_sSQLActRt, "sMCCatIDx = " & strParm(p_sMCCatIDx))
   Set p_oHistry = New Recordset
   p_oHistry.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   p_oHistry.Sort = "nAcctTerm DESC"
   p_oActRte.Sort = "nAcctTerm DESC"
   p_oActRte.MoveFirst
   
reCompareTerm:
   Do While p_oActRte.EOF = False
      If p_oHistry.EOF = False Then
         If CInt(p_oActRte("nAcctTerm")) > CInt(p_oHistry("nAcctTerm")) Then
            ' Modified > original
            If addNewRate() = False Then GoTo endWithRoll
         ElseIf CInt(p_oActRte("nAcctTerm")) = CInt(p_oHistry("nAcctTerm")) Then
            If CInt(p_oActRte("nAcctThru")) <> CInt(p_oHistry("nAcctThru")) Or _
                  p_oActRte("nFactorRt") <> p_oHistry("nFactorRt") Then
               If updateRate() = False Then GoTo endWithRoll
            End If
            p_oHistry.MoveNext
         Else
            ' Modified < Original
            Do
               If CInt(p_oActRte("nAcctTerm")) < CInt(p_oHistry("nAcctTerm")) Then
                  If removeRate() = False Then GoTo endWithRoll
               End If
               p_oHistry.MoveNext
            Loop Until p_oHistry.EOF
         End If
      Else
         If addNewRate() = False Then
            GoTo endWithRoll
         End If
      End If
      
      p_oActRte.MoveNext
   Loop
   
   With p_oHistry
      Do Until .EOF
         If removeRate() = False Then GoTo endWithRoll
         .MoveNext
      Loop
   End With
   
   If pbMasModified Then
      lsSQL = "UPDATE MC_Category SET" & _
                  "  nMiscChrg = " & p_nMiscChrg & _
                  ", nRebatesx = " & p_nRebatesx & _
                  ", nEndMrtgg = " & p_nEndMrtgg & _
               " WHERE sMCCatIDx = " & strParm(p_sMCCatIDx)
               
      If p_oAppDrivr.Execute(lsSQL, "MC_Category") <= 0 Then
         MsgBox "Unable to Update MC Category Info!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   p_oAppDrivr.CommitTrans
   
   pbMasModified = False
   pbDetModified = False
   pbRteModified = False
   
   SaveTransaction = True
   
endProc:
   Set p_oHistry = Nothing
   
   Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String

   lsProcName = "SearchDetail"
   On Error GoTo errProc
   SearchDetail = False

   Debug.Print p_oDetail.RecordCount
   If pbInitTran = False Then GoTo endProc

   If Row > (p_oDetail.RecordCount - 1) Then GoTo endProc
   If p_oDetail.AbsolutePosition <> Row + 1 Then
      p_oDetail.Move Row, adBookmarkFirst
   End If
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "smodelnme"
      getDetail Value, True
   End Select
   
   SearchDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchMaster(ByVal Index As Variant, Optional Value As Variant = "") As Boolean
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "smccatnme"
      Call getMaster(Value, True)
   End Select
   
   SearchMaster = True
End Function

Sub AddDetail()
   Dim lsProcName As String
   
   lsProcName = "addDetail"
   On Error GoTo errProc

   If p_oDetail("sModelIDx") = "" Then
      MsgBox "Inserting Empty Record is Not Allowed!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!", vbCritical, "Warning"
   End If
   
   p_oDetail.AddNew
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Sub AddTermRate()
   Dim lsProcName As String
   
   lsProcName = "addTermRate"
   On Error GoTo errProc

   If CInt(p_oActRte("nAcctTerm")) = 0 Then
      MsgBox "Inserting Empty Record is Not Allowed!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!", vbCritical, "Warning"
   End If
   
   p_oActRte.AddNew
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Function DeleteDetail(ByVal Row As Long) As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteDetail"
   On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If Row > (p_oDetail.RecordCount - 1) Then GoTo endProc
   
   p_oDetail.Move Row, adBookmarkFirst
   p_oDetail.Delete adAffectCurrent
   
   DeleteDetail = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function DeleteTermRate(ByVal Row As Long) As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTermRate"
   On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If Row > (p_oActRte.RecordCount - 1) Then GoTo endProc
   
   p_oActRte.Move Row, adBookmarkFirst
   p_oActRte.Delete adAffectCurrent
   
   DeleteTermRate = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub getMaster(ByVal lsValue As String, ByVal lbSearch As Boolean)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String

   lsProcName = "getMaster"
   On Error GoTo errProc
   
   If lsValue <> "" Then
      If lsValue = p_sMCCatNme Then GoTo endProc
      If lbSearch Then
         lsSQL = "sMCCatNme LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "sMCCatNme = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  sMCCatIDx" & _
               ", sMCCatNme" & _
               ", nMiscChrg" & _
               ", nRebatesx" & _
               ", nEndMrtgg" & _
            " FROM MC_Category" & _
            IIf(lsSQL <> Empty, " WHERE " & lsSQL, "")
            
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      GoTo endWithClear
   ElseIf loRS.RecordCount > 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, loRS _
                           , "sMCCatIDx»sMCCatNme" _
                           , "Code»Category Name")

      If lsSelect <> "" Then
         lasSelect = Split(lsSelect, "»")
         
         loRS.MoveFirst
         loRS.Find "sMCCatIDx = " & strParm(lasSelect(0)), 0, adSearchForward
         Debug.Print loRS.EOF
      End If
   End If
   
   If pbDetModified Or pbRteModified Then
      If p_sMCCatIDx <> loRS("sMCCatIDx") Then
         If MsgBox("Some of the Info were modified!" & vbCrLf & _
                  "Loading New Set of MC Model will DISCARD modified Information!" & vbCrLf & vbCrLf & _
                  "CONTINUE ANYWAY?", vbQuestion + vbYesNo, "Confirm") <> vbYes Then
            GoTo endProc
         End If
      End If
   End If
   
   p_sMCCatIDx = loRS("sMCCatIDx")
   p_sMCCatNme = loRS("sMCCatNme")
   p_nMiscChrg = IFNull(loRS("nMiscChrg"), 0)
   p_nRebatesx = IFNull(loRS("nRebatesx"), 0)
   p_nEndMrtgg = IFNull(loRS("nEndMrtgg"), 0)
   
   Call loadModel
   Call loadTerm
   
endProc:
   RaiseEvent MasterRetrieved(1)
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   RaiseEvent MasterRetrieved(5)
   Exit Sub
endWithClear:
   p_sMCCatIDx = ""
   p_sMCCatNme = ""
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Sub loadModel()
   Dim lsProcName As String

   lsProcName = "loadModel"
   On Error GoTo errProc
   
   Set p_oDetail = New Recordset
   With p_oDetail
      .Open AddCondition(p_sSQLModel, "a.sMCCatIDx = " & strParm(p_sMCCatIDx)), _
               p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      
      ' if no record exist, add empty rec
      Set .ActiveConnection = Nothing
      If .EOF Then p_oDetail.AddNew
   End With
   
   pbDetModified = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub loadTerm()
   Dim lsProcName As String

   lsProcName = "loadTerm"
   On Error GoTo errProc
   
   Set p_oActRte = New Recordset
   With p_oActRte
      .Open AddCondition(p_sSQLActRt, "sMCCatIDx = " & strParm(p_sMCCatIDx)), _
               p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      
      ' if no record exist, add empty rec
      Set .ActiveConnection = Nothing
      If .EOF Then p_oActRte.AddNew
   End With
   
   pbRteModified = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub getDetail(ByVal lsValue As String, _
                        ByVal lbSearch As Boolean)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String

   lsProcName = "getDetail"
   On Error GoTo errProc
   
   If lsValue <> "" Then
      If lsValue = p_oDetail("sModelNme") Then GoTo endProc
      If lbSearch Then
         lsSelect = "a.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSelect = "a.sModelNme = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", a.sModelCde" & _
               ", c.sBrandNme" & _
               ", b.nSelPrice" & _
               ", b.dInsPrice" & _
               ", b.nMinDownx" & _
               ", b.sModified" & _
               ", b.dModified" & _
               ", b.sMCCatIDx" & _
               ", b.cRecdStat" & _
            " FROM MC_Model a" & _
               " LEFT JOIN MC_Model_Price b" & _
                  " ON a.sModelIDx = b.sModelIDx" & _
               ", Brand c" & _
            " WHERE a.sBrandIDx = c.sBrandIDx" & _
               " AND (b.sMCCatIDx IS NULL" & _
                  " OR b.sMCCatIDx <> " & strParm(p_sMCCatIDx) & _
                  " OR ( b.sMCCatIDx = " & strParm(p_sMCCatIDx) & _
                     " AND b.cRecdStat = " & strParm(xeRecStateInactive) & " ) )" & _
            " ORDER BY a.sModelNme"
            
   If lsSelect <> "" Then
      lsSQL = AddCondition(lsSQL, lsSelect)
   End If
            
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      GoTo endWithClear
   ElseIf loRS.RecordCount > 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, loRS _
                           , "sModelIDx»sModelNme»sBrandNme" _
                           , "Code»Model Name»Brand")

      If lsSelect = "" Then
         If lbSearch Then
            GoTo endProc
         Else
            GoTo endWithClear
         End If
      Else
         lasSelect = Split(lsSelect, "»")
         
         loRS.MoveFirst
         loRS.Find "sModelIDx = " & strParm(lasSelect(0)), 1, adSearchForward
      End If
   End If
   
   p_oDetail("sModelIDx") = loRS("sModelIDx")
   p_oDetail("sModelNme") = loRS("sModelNme")
   p_oDetail("sModelCde") = loRS("sModelCde")
   p_oDetail("sBrandNme") = loRS("sBrandNme")
   If IsNull(loRS("nSelPrice")) Then
      p_oDetail("nSelPrice") = 0#
      p_oDetail("nMinDownx") = 0#
      p_oDetail("dInsPrice") = p_dInsPrice
      p_oDetail("sMCCatIDx") = ""
      p_oDetail("cRecdStat") = xeRecStateActive
   Else
      p_oDetail("nSelPrice") = loRS("nSelPrice")
      p_oDetail("nMinDownx") = loRS("nMinDownx")
      p_oDetail("dInsPrice") = loRS("dInsPrice")
      p_oDetail("sMCCatIDx") = loRS("sMCCatIDx")
      p_oDetail("cRecdStat") = xeRecStateUnknown
   End If
   
endProc:
   RaiseEvent DetailRetrieved
   Exit Sub
endWithClear:
   p_oDetail("sModelIDx") = ""
   p_oDetail("sModelNme") = ""
   p_oDetail("sModelCde") = ""
   p_oDetail("nSelPrice") = 0#
   p_oDetail("nMinDownx") = 0#
   p_oDetail("sMCCatIDx") = ""
   p_oDetail("cRecdStat") = xeRecStateInactive
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Function addNewModel() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "addNewModel"
   On Error GoTo errProc
   
   'From other model
   If p_oDetail("cRecdStat") = xeRecStateUnknown Then
      If createNotice(p_oDetail("sModelIDx"), p_oDetail("sMCCatIDx"), p_sMCCatIDx) Then
         addNewModel = updateMCCategory
      End If
      GoTo endProc
   Else
      ' create a record that will inform installment price update
      '  of a new model
      If createNotice(p_oDetail("sModelIDx"), "", p_sMCCatIDx) = False Then GoTo endProc
   End If
   
   lsSQL = "INSERT INTO MC_Model_Price SET" & _
               "  sModelIDx = " & strParm(p_oDetail("sModelIDx")) & _
               ", nSelPrice = " & p_oDetail("nSelPrice") & _
               ", nLastPrce = " & p_oDetail("nSelPrice") & _
               ", nDealrPrc = " & p_oDetail("nSelPrice") & _
               ", nMinDownx = " & p_oDetail("nMinDownx") & _
               ", sMCCatIDx = " & strParm(p_sMCCatIDx) & _
               ", sApproved = " & strParm(p_oAppDrivr.UserID) & _
               ", dInsPrice = " & dateParm(p_dInsPrice) & _
               ", cRecdStat = " & strParm(xeRecStateActive) & _
               ", sPrepared = " & strParm(p_oAppDrivr.UserID) & _
               ", dPrepared = " & dateParm(p_oAppDrivr.SysDate) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate)
   
   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price") <= 0 Then
      MsgBox "Unable to Insert New Record!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   addNewModel = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function updateModelPrice() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "updateModelPrice"
   On Error GoTo errProc
   
   ' first move the existing record to history table
   lsSQL = "INSERT INTO MC_Model_Price_History " & _
               "( sModelIDx" & _
               ", dPricexxx" & _
               ", nPurPrice" & _
               ", nSelPrice" & _
               ", nLastPrce" & _
               ", nDealrPrc" & _
               ", nMinDownx" & _
               ", sMCCatIDx" & _
               ", sModified" & _
               ", dModified )" & _
               " ( SELECT" & _
                  "  sModelIDx" & _
                  ", dInsPrice" & _
                  ", nPurPrice" & _
                  ", nSelPrice" & _
                  ", nLastPrce" & _
                  ", nDealrPrc" & _
                  ", nMinDownx" & _
                  ", sMCCatIDx" & _
                  ", sModified" & _
                  ", dModified" & _
               " FROM MC_Model_Price" & _
               " WHERE sModelIDx = " & strParm(p_oHistry("sModelIDx")) & " )"

   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price_History") <= 0 Then
      MsgBox "Unable to Insert Record to History!", vbCritical, "Warning"
      GoTo endProc
   End If
            
   ' now update the current pricelist
   lsSQL = "UPDATE MC_Model_Price SET" & _
               "  dInsPrice = " & dateParm(p_dInsPrice) & _
               ", nSelPrice = " & p_oDetail("nSelPrice") & _
               ", nMinDownx = " & p_oDetail("nMinDownx") & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate) & _
            " WHERE sModelIDx = " & strParm(p_oDetail("sModelIDx"))
   
   Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price") <= 0 Then
      MsgBox "Unable to Update Price to Info!", vbCritical, "Warning"
      GoTo endProc
   End If

   updateModelPrice = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function removeModel() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "removeModel"
   On Error GoTo errProc
   
   ' first move the existing record to history table
   lsSQL = "INSERT INTO MC_Model_Price_History " & _
               "( sModelIDx" & _
               ", dPricexxx" & _
               ", nPurPrice" & _
               ", nSelPrice" & _
               ", nLastPrce" & _
               ", nDealrPrc" & _
               ", nMinDownx" & _
               ", sMCCatIDx" & _
               ", sModified" & _
               ", dModified )" & _
               " ( SELECT" & _
                  "  sModelIDx" & _
                  ", dInsPrice" & _
                  ", nPurPrice" & _
                  ", nSelPrice" & _
                  ", nLastPrce" & _
                  ", nDealrPrc" & _
                  ", nMinDownx" & _
                  ", sMCCatIDx" & _
                  ", sModified" & _
                  ", dModified" & _
               " FROM MC_Model_Price" & _
               " WHERE sModelIDx = " & strParm(p_oHistry("sModelIDx")) & " )"

   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price_History") <= 0 Then
      MsgBox "Unable to Insert Record to History!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   ' now update the current pricelist
   lsSQL = "UPDATE MC_Model_Price SET" & _
               "  cRecdStat = " & strParm(xeRecStateInactive) & _
               ", sMCCatIDx = " & strParm("") & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate) & _
            " WHERE sModelIDx = " & strParm(p_oHistry("sModelIDx"))
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price") <= 0 Then
      MsgBox "Unable to Delete Price Info!", vbCritical, "Warning"
      GoTo endProc
   End If

   removeModel = createNotice(p_oDetail("sModelIDx"), p_oDetail("sMCCatIDx"), "")

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function createNotice(lsModelIDx As String, lsPrvCatIDx As String, lsMCCatIDx As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "createNotice"
   On Error GoTo errProc
   
   ' now update the current pricelist
   lsSQL = "INSERT INTO MC_Category_Changes SET" & _
               "  sModelIDx = " & strParm(lsModelIDx) & _
               ", sPrvCatID = " & strParm(lsPrvCatIDx) & _
               ", sPrsCatID = " & strParm(lsMCCatIDx) & _
               ", dTransact = " & dateParm(p_dInsPrice) & _
               ", cChangeFr = " & strParm(pxeChangeFrom) & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate)

   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price") <= 0 Then GoTo endProc

   createNotice = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function updateMCCategory() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "updateMCCategory"
   On Error GoTo errProc
   Debug.Print lsProcName
   
   ' first move the existing record to history table
   lsSQL = "INSERT INTO MC_Model_Price_History " & _
               "( sModelIDx" & _
               ", dPricexxx" & _
               ", nPurPrice" & _
               ", nSelPrice" & _
               ", nLastPrce" & _
               ", nDealrPrc" & _
               ", nMinDownx" & _
               ", sMCCatIDx" & _
               ", sModified" & _
               ", dModified )" & _
               " ( SELECT" & _
                  "  sModelIDx" & _
                  ", dInsPrice" & _
                  ", nPurPrice" & _
                  ", nSelPrice" & _
                  ", nLastPrce" & _
                  ", nDealrPrc" & _
                  ", nMinDownx" & _
                  ", sMCCatIDx" & _
                  ", sModified" & _
                  ", dModified" & _
               " FROM MC_Model_Price" & _
               " WHERE sModelIDx = " & strParm(p_oDetail("sModelIDx"))
   If Not IsNull(p_oDetail("dInsPrice")) Then
      lsSQL = lsSQL & _
                  " AND dInsPrice = " & dateParm(p_oDetail("dInsPrice")) & ")"
   End If
Debug.Print lsSQL
   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price_History") <= 0 Then
      MsgBox "Unable to Insert Record to History!", vbCritical, "Warning"
      GoTo endProc
   End If
            
   ' now update the current pricelist
   lsSQL = "UPDATE MC_Model_Price SET" & _
               "  sModelIDx = " & strParm(p_oDetail("sModelIDx")) & _
               ", dInsPrice = " & dateParm(p_dInsPrice) & _
               ", nSelPrice = " & p_oDetail("nSelPrice") & _
               ", nMinDownx = " & p_oDetail("nMinDownx") & _
               ", sMCCatIDx = " & strParm(p_sMCCatIDx) & _
               ", cRecdStat = " & strParm(xeRecStateActive) & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate) & _
            " WHERE sModelIDx = " & strParm(p_oDetail("sModelIDx"))
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Model_Price") <= 0 Then
      MsgBox "Unable to Update Price to Info!", vbCritical, "Warning"
      GoTo endProc
   End If

   updateMCCategory = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function addNewRate() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "addNewRate"
   On Error GoTo errProc
   
   lsSQL = "INSERT INTO MC_Term_Category SET" & _
               "  sMCCatIDx = " & strParm(p_sMCCatIDx) & _
               ", nAcctTerm = " & p_oActRte("nAcctTerm") & _
               ", nAcctThru = " & p_oActRte("nAcctThru") & _
               ", nFactorRt = " & p_oActRte("nFactorRt") & _
               ", sApproved = " & strParm(p_oAppDrivr.UserID) & _
               ", dPricexxx = " & dateParm(p_dInsPrice) & _
               ", sPrepared = " & strParm(p_oAppDrivr.UserID) & _
               ", dPrepared = " & dateParm(p_oAppDrivr.SysDate) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate)
   
   If p_oAppDrivr.Execute(lsSQL, "MC_Term_Category") <= 0 Then
      MsgBox "Unable to Insert New Record!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   addNewRate = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function updateRate() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "updateRate"
   On Error GoTo errProc
   
   ' first move the existing record to history table
   lsSQL = "INSERT INTO MC_Term_Category_History " & _
               "( sMCCatIDx" & _
               ", nAcctTerm" & _
               ", nAcctThru" & _
               ", nFactorRt" & _
               ", dPricexxx" & _
               ", sModified" & _
               ", dModified )" & _
               " ( SELECT" & _
                  "  sMCCatIDx" & _
                  ", nAcctTerm" & _
                  ", nAcctThru" & _
                  ", nFactorRt" & _
                  ", dPricexxx" & _
                  ", sModified" & _
                  ", dModified" & _
               " FROM MC_Term_Category" & _
               " WHERE sMCCatIDx = " & strParm(p_oHistry("sMCCatIDx")) & _
                  " AND nAcctTerm = " & p_oHistry("nAcctTerm") & " )"

   If p_oAppDrivr.Execute(lsSQL, "MC_Term_Category_History") <= 0 Then
      MsgBox "Unable to Insert Record to History!", vbCritical, "Warning"
      GoTo endProc
   End If
            
   ' now update the current rate
   lsSQL = "UPDATE MC_Term_Category SET" & _
               "  nAcctThru = " & p_oActRte("nAcctThru") & _
               ", nFactorRt = " & p_oActRte("nFactorRt") & _
               ", sApproved = " & strParm(p_oAppDrivr.UserID) & _
               ", dPricexxx = " & dateParm(p_dInsPrice) & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.SysDate) & _
            " WHERE sMCCatIDx = " & strParm(p_sMCCatIDx) & _
               " AND nAcctTerm = " & p_oActRte("nAcctTerm")
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Term_Category") <= 0 Then
      MsgBox "Unable to Update Price to Info!", vbCritical, "Warning"
      GoTo endProc
   End If

   updateRate = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function removeRate() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "removeModel"
   On Error GoTo errProc
   
   ' first move the existing record to history table
   lsSQL = "INSERT INTO MC_Term_Category_History " & _
               "( sMCCatIDx" & _
               ", nAcctTerm" & _
               ", nAcctThru" & _
               ", nFactorRt" & _
               ", dPricexxx" & _
               ", sModified" & _
               ", dModified )" & _
               " ( SELECT" & _
                  "  sMCCatIDx" & _
                  ", nAcctTerm" & _
                  ", nAcctThru" & _
                  ", nFactorRt" & _
                  ", dPricexxx" & _
                  ", sModified" & _
                  ", dModified )" & _
               " FROM MC_Term_Category" & _
               " WHERE sMCCatIDx = " & strParm(p_oHistry("sMCCatIDx")) & _
                  " AND nAcctTerm = " & p_oHistry("nAcctTerm") & " )"

   If p_oAppDrivr.Execute(lsSQL, "MC_Term_Category_History") <= 0 Then
      MsgBox "Unable to Insert Record to History!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   ' now update the current pricelist
   lsSQL = "DELETE FROM MC_Term_Category" & _
            " WHERE sMCCatIDx = " & strParm(p_oHistry("sMCCatIDx")) & _
               " AND nAcctTerm = " & p_oHistry("nAcctTerm")
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Term_Category") <= 0 Then
      MsgBox "Unable to Delete Term Info!", vbCritical, "Warning"
      GoTo endProc
   End If

   removeRate = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

