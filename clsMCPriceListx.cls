VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC Price List
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
'  XerSys [ 07/06/2009 08:43 am ]
'     Start coding this object...
'  XerSys [ 09/30/2011 06:42 am ]
'     Integrate end of mortgage fee.
'  XerSys [ 05/08/2013 11:23 am ]
'     Insert searching of model by id
'  XerSys [ 05/22/2013 05:11 pm ]
'     Short term payments don't follow the standard computation
'  XerSys [ 08/16/2013 04:08 pm ]
'     Include add on amount to loaned motorcycle such as side car
'  Mac  [ 07/30/2020 04:33 pm ]
'     Incorporate computation for REPO units
'        added MCSerialID Property
'        added getSerial Function
'
'     NOTE:
'        if unit was a BRAND NEW use MCModelID Let Property to initialize
'        if unit was a REPO use MCSerialID Let Property to initialize
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMCPriceList"

Private p_oAppDrivr As clsAppDriver
Private p_oCashPrice As Recordset
Private p_oInsPrice As Recordset

Private p_sMCCatIDx As String
Private p_sMCCatNme As String
Private p_sModelIDx As String
Private p_sModelNme As String

Private p_nSelPrice As Double
Private p_nLastPrce As Double
Private p_nRebatesx As Double
Private p_nMiscChrg As Double
Private p_nMinDownx As Double
Private p_nEndMrtgg As Double

Private p_nAddPurc As Double

Private p_sSQLCash As String
Private p_sSQLInst As String
Private p_bExactNm As Boolean
Private p_bByCode As Boolean

Private pbInitTran As Boolean

Private p_sSerialID As String
Private p_cSoldStat As String

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get CashPrice(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   With p_oCashPrice
      If Row > (.RecordCount - 1) Then Exit Property
      .Move Row, adBookmarkFirst
      
      CashPrice = .Fields(Index)
   End With
End Property

Property Let CashPrice(ByVal Row As Long, Index As Variant, Value As Variant)
End Property

Property Get AddtlLoan() As Double
   AddtlLoan = p_nAddPurc
End Property

Property Let AddtlLoan(ByVal LoanAmt As Double)
   p_nAddPurc = LoanAmt
End Property

Property Get InstallmentPrice(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   With p_oInsPrice
      If Row > (.RecordCount - 1) Then Exit Property
      .Move Row, adBookmarkFirst
      
      If Index = "dInsPrice" Then
         InstallmentPrice = IIf(.Fields("dInsPrice") > .Fields("xInsPrice"), .Fields("dInsPrice"), .Fields("xInsPrice"))
      Else
         InstallmentPrice = .Fields(Index)
      End If
   End With
End Property

Property Let InstallmentPrice(ByVal Row As Long, Index As Variant, Value As Variant)
End Property

Property Get MCCategory() As String
   MCCategory = p_sMCCatNme
End Property

Property Let MCCategory(ByVal Value As String)
   Call getCategory(Value)
End Property

Property Get MCSerialID() As String
   MCSerialID = p_sSerialID
End Property

Property Let MCSerialID(ByVal fsSerialID As String)
   Call getSerial(fsSerialID)
End Property

Property Get MCModelID() As String
   MCModelID = p_sModelIDx
End Property

Property Let MCModelID(ByVal Value As String)
   p_bByCode = True
   Call getModel(Value)
End Property

Property Get MCModel() As String
   MCModel = p_sModelNme
End Property

Property Let MCModel(ByVal Value As String)
   p_bByCode = False
   Call getModel(Value)
End Property

Property Get Rebate() As Double
   Rebate = p_nRebatesx
End Property

Property Let Rebate(Value As Double)
End Property

Property Get MinimumDown() As Double
   MinimumDown = p_nMinDownx
End Property

Property Let MinimumDown(Value As Double)
End Property

Property Get MiscCharge() As Double
   MiscCharge = p_nMiscChrg
End Property

Property Let MiscCharge(Value As Double)
End Property

Property Get EndMortgage() As Double
   EndMortgage = p_nEndMrtgg
End Property

Property Let EndMortgage(Value As Double)
End Property

Property Get LastPrice() As Double
   LastPrice = p_nLastPrce
End Property

Property Let LastPrice(Value As Double)
End Property

Property Get SelPrice() As Double
   SelPrice = p_nSelPrice
End Property

Property Let SelPrice(Value As Double)
End Property

Property Get CashPriceCount() As Long
   If pbInitTran = False Then Exit Property
   
   CashPriceCount = p_oCashPrice.RecordCount
End Property

Property Get InstallmentPriceCount() As Long
   If pbInitTran = False Then Exit Property
   
   InstallmentPriceCount = p_oInsPrice.RecordCount
End Property

Property Get CashLatestDate() As Date
   Dim ldLatest As Date
   
   CashLatestDate = CDate("2001-01-01")
   With p_oCashPrice
      If .RecordCount > 1 Then
         .MoveFirst
         ldLatest = IFNull(.Fields("dPricexxx"), "2001-01-01")
         Do While .EOF = False
            If ldLatest < .Fields("dPricexxx") Then ldLatest = .Fields("dPricexxx")
            .MoveNext
         Loop
         CashLatestDate = ldLatest
      End If
   End With
End Property

Property Get InstallmentLatestDate() As Date
   Dim ldLatest As Date
   
   InstallmentLatestDate = CDate("2001/01/01")
   With p_oInsPrice
      If .RecordCount > 1 Then
         .MoveFirst
         ldLatest = IIf(.Fields("dInsPrice") > .Fields("xInsPrice"), .Fields("dInsPrice"), .Fields("xInsPrice"))
         Do While .EOF = False
            If ldLatest < .Fields("dInsPrice") Then ldLatest = .Fields("dInsPrice")
            If ldLatest < .Fields("xInsPrice") Then ldLatest = .Fields("xInsPrice")
            .MoveNext
         Loop
         InstallmentLatestDate = ldLatest
      End If
   End With
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsAddConditon As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'she 2022-08-06
   'as per request by GMA. to include all list to 2H branches
   
   Select Case p_oAppDrivr.BranchCode
   Case "M2H1", "M187", "M186", "M102", "M0R1", "M085"
      lsAddConditon = ""
   Case Else
      lsAddConditon = " AND b.cRecdStat = " & strParm(xeRecStateActive)
   End Select
   
   p_sSQLCash = "SELECT" & _
                     "  c.sMCCatNme" & _
                     ", b.sModelNme" & _
                     ", d.sBrandNme" & _
                     ", a.nSelPrice" & _
                     ", a.nLastPrce" & _
                     ", a.nDealrPrc" & _
                     ", a.dPricexxx" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                     ", MC_Category c" & _
                     ", Brand d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sMCCatIDx = c.sMCCatIDx" & _
                     " AND b.sBrandIDx = d.sBrandIDx" & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     lsAddConditon & _
                  " ORDER BY c.sMCCatNme, b.sModelNme"
   
   p_sSQLInst = "SELECT" & _
                     "  c.sMCCatNme" & _
                     ", b.sModelNme" & _
                     ", a.nSelPrice" & _
                     ", a.nMinDownx" & _
                     ", d.nAcctThru" & _
                     ", d.nFactorRt" & _
                     ", a.dInsPrice" & _
                     ", d.dPricexxx xInsPrice" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                     ", MC_Category c" & _
                     ", MC_Term_Category d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sMCCatIDx = c.sMCCatIDx" & _
                     " AND a.sMCCatIDx = d.sMCCatIDx" & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                  " ORDER BY c.sMCCatNme, b.sModelNme"
                  
   Set p_oCashPrice = New Recordset
   With p_oCashPrice
      .Open AddCondition(p_sSQLCash, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
   End With
   
   Set p_oInsPrice = New Recordset
   With p_oInsPrice
      .Open AddCondition(p_sSQLInst, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
   End With
   
   pbInitTran = True
   InitTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Sub LoadCashPrice()
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "LoadCashPrice"
   On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   If p_sMCCatIDx <> "" Then
      lsSQL = AddCondition(p_sSQLCash, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
   ElseIf p_sModelIDx <> "" Then
      lsSQL = AddCondition(p_sSQLCash, "a.sModelIDx = " & strParm(p_sModelIDx))
   Else
      lsSQL = p_sSQLCash
   End If
   
   Debug.Print lsSQL
   Set p_oCashPrice = New Recordset
   With p_oCashPrice
      Debug.Print lsSQL
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
   End With
   
   p_bExactNm = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Sub LoadInstallmentPrice()
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "LoadInstallmentPrice"
   On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   lsSQL = p_sSQLInst
   If p_sMCCatIDx <> "" Then
      lsSQL = AddCondition(p_sSQLInst, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
   End If
   Debug.Print lsSQL
   Set p_oInsPrice = New Recordset
   With p_oInsPrice
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
   End With
   
   p_bExactNm = True
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Function getMonthly(ByVal DownPayment As Double, Term As Integer) As Double
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnSelPrice As Double
   Dim lnMinRatex As Double
   
   lsProcName = "getMonthly"
   On Error GoTo errProc
   
   Set loRS = New Recordset
   
   If p_sModelIDx = "" Then
      MsgBox "No Model has been Selected!" & vbCrLf & _
               "Please select a model first then try again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lnSelPrice = 0#
   
   If p_cSoldStat = "1" Then
      lsSQL = "SELECT" & _
                  "  a.sSerialID" & _
                  ", a.sModelIDx" & _
                  ", b.nPOwnSRPx" & _
               " FROM MC_Serial a" & _
                  " LEFT JOIN MC_Model d" & _
                     " ON a.sModelIDx = d.sModelIDx" & _
                  " LEFT JOIN MC_Serial_Classification_New b" & _
                     " ON a.sSerialID = b.sSerialID" & _
               " WHERE a.sSerialID = " & strParm(p_sSerialID) & _
                  " AND a.cSoldStat = '1'"
   
      Debug.Print lsSQL
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      Set loRS.ActiveConnection = Nothing
      
      If Not loRS.EOF Then
         lnSelPrice = IFNull(loRS("nPOwnSRPx"), 0#)
      End If
   End If
   'end - mac 2020.09.05
   
   lsSQL = "SELECT" & _
               "  a.nSelPrice" & _
               ", a.nMinDownx" & _
               ", b.nMiscChrg" & _
               ", b.nRebatesx" & _
               ", b.nEndMrtgg" & _
               ", c.nAcctThru" & _
               ", c.nFactorRt" & _
            " FROM MC_Model_Price a" & _
               ", MC_Category b" & _
               ", MC_Term_Category c" & _
            " WHERE a.sMCCatIDx = b.sMCCatIDx" & _
               " AND a.sMCCatIDx = c.sMCCatIDx" & _
               " AND a.sModelIDx = " & strParm(p_sModelIDx) & _
               " AND c.nAcctThru = " & Term
   
   If loRS.State = adStateOpen Then loRS.Close
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   If loRS.EOF Then GoTo endProc
   
   If p_cSoldStat = "0" Then lnSelPrice = loRS("nSelPrice")
   
   ' XerSys - 2013-08-19
   '  Update Computation of MA. The additional loan such as sidecar
   '     must be incorporated at the SRP
   If Term < 4 Then
      ' XerSys 2013-05-22
      '  Three months or less don't follow the standard computation
      'getMonthly = Round(((loRS("nSelPrice") + p_nAddPurc - (DownPayment)) _
      '               * loRS("nFactorRt") / Term), 0)
      
      getMonthly = Round(((lnSelPrice + p_nAddPurc - (DownPayment)) _
                     * loRS("nFactorRt") / Term), 0)
   Else
      'getMonthly = Round(((loRS("nSelPrice") + p_nAddPurc - DownPayment + loRS("nMiscChrg")) _
      '               * loRS("nFactorRt") / Term) + loRS("nRebatesx") + (loRS("nEndMrtgg") / Term), 0)
                     
      getMonthly = Round(((lnSelPrice + p_nAddPurc - DownPayment + loRS("nMiscChrg")) _
                     * loRS("nFactorRt") / Term) + loRS("nRebatesx") + (loRS("nEndMrtgg") / Term), 0)
   End If
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & DownPayment _
                        & ", " & Term & " )"
End Function

Private Sub getCategory(ByVal lsValue As String)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String

   lsProcName = "getCategory"
   On Error GoTo errProc
   
   If lsValue = p_sMCCatNme Then GoTo endProc
  
   lsSQL = "SELECT" & _
               "  sMCCatIDx" & _
               ", sMCCatNme" & _
            " FROM MC_Category" & _
            " WHERE sMCCatNme LIKE " & strParm(Trim(lsValue) & "%")
            
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
      End If
   End If
   
   p_sMCCatIDx = loRS("sMCCatIDx")
   p_sMCCatNme = loRS("sMCCatNme")
   
endProc:
   Exit Sub
endWithClear:
   p_sMCCatIDx = ""
   p_sMCCatNme = ""
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue & " )"
End Sub

'mac 2020.07.31
'  created this procedure
'mac 2022.08.22
'  revise it for the new mc serial classification logic
Private Sub getSerial(ByVal fsSerialID As String)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   
   Dim lnMinRatex As Double
   Dim lnPOwnSRPx As Double

   lsProcName = "getSerial"
   On Error GoTo errProc
   
   Set loRS = New Recordset
               
   lsSQL = "SELECT" & _
               "  a.sSerialID" & _
               ", a.sModelIDx" & _
               ", b.nPOwnSRPx" & _
               ", b.nMinDownx" & _
            " FROM MC_Serial a" & _
               " LEFT JOIN MC_Model d" & _
                  " ON a.sModelIDx = d.sModelIDx" & _
               " LEFT JOIN MC_Serial_Classification_New b" & _
                  " ON a.sSerialID = b.sSerialID" & _
            " WHERE a.sSerialID = " & strParm(fsSerialID) & _
               " AND a.cSoldStat = '1'"
   
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then GoTo endWithClear
   
   p_nMinDownx = IFNull(loRS("nMinDownx"), 0#)
   
   If IsNull(loRS("nPOwnSRPx")) Then GoTo endWithClear
   
   lnPOwnSRPx = loRS("nPOwnSRPx")
   
   lsSQL = "a.sModelIDx = " & strParm(loRS("sModelIDx"))
   
   lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", b.nMinDownx" & _
               ", c.nRebatesx" & _
               ", c.nMiscChrg" & _
               ", c.nEndMrtgg" & _
               ", b.nSelPrice" & _
               ", b.nLastPrce" & _
               ", a.cHotItemx" & _
            " FROM MC_Model a" & _
               ", MC_Model_Price b" & _
               ", MC_Category c" & _
            " WHERE a.sModelIDx = b.sModelIDx" & _
               " AND b.sMCCatIDx = c.sMCCatIDx" & _
               IIf(lsSQL <> Empty, " AND " & lsSQL, "")
            
   Debug.Print lsSQL
   If loRS.State = adStateOpen Then loRS.Close
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      GoTo endWithClear
   ElseIf loRS.RecordCount > 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, loRS _
                           , "sModelIDx»sModelNme" _
                           , "Code»Model Name")

      If lsSelect <> "" Then
         lasSelect = Split(lsSelect, "»")
         
         loRS.MoveFirst
         loRS.Find "sModelIDx = " & strParm(lasSelect(0)), 0, adSearchForward
      End If
   End If
   
   p_sModelIDx = loRS("sModelIDx")
   p_sModelNme = loRS("sModelNme")
   p_nRebatesx = loRS("nRebatesx")
   p_nMiscChrg = loRS("nMiscChrg")
   p_nEndMrtgg = loRS("nEndMrtgg")
   p_nSelPrice = lnPOwnSRPx 'set the SRP as from the classification
   p_nLastPrce = lnPOwnSRPx 'set the last price as the SRP
endProc:
   ' XerSys - 2013-08-16
   '  Reset the additional purchase for every model resets
   p_nAddPurc = 0
     
   p_cSoldStat = "1"
   p_sSerialID = fsSerialID
   
   Exit Sub
endWithClear:
   p_sModelIDx = ""
   p_sModelNme = ""
   p_nRebatesx = 0
   p_nMiscChrg = 0
   p_nEndMrtgg = 0
   
   p_cSoldStat = "0"
   p_sSerialID = ""
   
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & fsSerialID & " )"
End Sub

Private Sub getModel(ByVal lsValue As String)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String

   lsProcName = "getModel"
   On Error GoTo errProc
   
   If p_bByCode Then
      'If lsValue = p_sModelIDx Then GoTo endProc
      lsSQL = "a.sModelIDx = " & strParm(Trim(lsValue))
   Else
      'If lsValue = p_sModelNme Then GoTo endProc
      If p_bExactNm Then
         lsSQL = "a.sModelNme = " & strParm(Trim(lsValue))
      Else
         lsSQL = "a.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
      End If
   End If
   
   lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", b.nMinDownx" & _
               ", c.nRebatesx" & _
               ", c.nMiscChrg" & _
               ", c.nEndMrtgg" & _
               ", b.nSelPrice" & _
               ", b.nLastPrce" & _
            " FROM MC_Model a" & _
               ", MC_Model_Price b" & _
               ", MC_Category c" & _
            " WHERE a.sModelIDx = b.sModelIDx" & _
               " AND b.sMCCatIDx = c.sMCCatIDx" & _
               IIf(lsSQL <> Empty, " AND " & lsSQL, "") & _
            " AND a.cRecdStat = " & strParm(xeRecStateActive)
            
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.EOF Then
      GoTo endWithClear
   ElseIf loRS.RecordCount > 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, loRS _
                           , "sModelIDx»sModelNme" _
                           , "Code»Model Name")

      If lsSelect <> "" Then
         lasSelect = Split(lsSelect, "»")
         
         loRS.MoveFirst
         loRS.Find "sModelIDx = " & strParm(lasSelect(0)), 0, adSearchForward
      Else
         GoTo endWithClear
      End If
   End If
   
   p_sModelIDx = loRS("sModelIDx")
   p_sModelNme = loRS("sModelNme")
   p_nRebatesx = loRS("nRebatesx")
   p_nMiscChrg = loRS("nMiscChrg")
   p_nEndMrtgg = loRS("nEndMrtgg")
   p_nMinDownx = IIf(IsNull(loRS("nMinDownx")), 0, loRS("nMinDownx"))
   p_nSelPrice = loRS("nSelPrice")
   p_nLastPrce = loRS("nLastPrce")
   
endProc:
   ' XerSys - 2013-08-16
   '  Reset the additional purchase for every model resets
   p_nAddPurc = 0
   
   'mac 2020.07.31
   p_cSoldStat = "0"
   p_sSerialID = ""
   
   Exit Sub
endWithClear:
   p_sModelIDx = ""
   p_sModelNme = ""
   p_nRebatesx = 0
   p_nMiscChrg = 0
   p_nEndMrtgg = 0
   
   'mac 2020.07.31
   p_cSoldStat = "0"
   p_sSerialID = ""
   
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue & " )"
End Sub

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

'Private Sub getSerial(ByVal fsSerialID As String)
'   Dim loRS As Recordset
'   Dim lsProcName As String
'   Dim lsSelect As String
'   Dim lasSelect() As String
'   Dim lsSQL As String
'
'   Dim lnMinRatex As Double
'
'   lsProcName = "getSerial"
'   On Error GoTo errProc
'
'   Set loRS = New Recordset
'
'   lsSQL = "SELECT" & _
'               "  a.sSerialID" & _
'               ", a.sModelIDx" & _
'               ", b.sClassIDx" & _
'               ", c.nMinRatex" & _
'               ", c.nMaxRatex" & _
'               ", b.nMinDownx" & _
'               ", b.nSpclRate" & _
'            " FROM MC_Serial a" & _
'               " LEFT JOIN MC_Model d" & _
'                  " ON a.sModelIDx = d.sModelIDx" & _
'               " LEFT JOIN MC_Serial_Classification b" & _
'                  " ON a.sSerialID = b.sSerialID" & _
'               " LEFT JOIN MC_2H_Classification c" & _
'                  " ON b.sClassIDx = c.sClassIDx" & _
'                     " AND c.cHotModel = d.cHotItemx" & _
'                     " AND c.cRecdStat = '1'" & _
'            " WHERE a.sSerialID = " & strParm(fsSerialID) & _
'               " AND a.cSoldStat = '1'"
'
'   Debug.Print lsSQL
'   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'   Set loRS.ActiveConnection = Nothing
'
'   If loRS.EOF Then GoTo endWithClear
'   If IsNull(loRS("sClassIDx")) Then GoTo endWithClear
'
'   lnMinRatex = 0#
'
'   If IFNull(loRS("nSpclRate"), 0#) <= 0# Then
'      'minimum rate
'      lnMinRatex = IFNull(loRS("nMinRatex"), 0)
'   Else
'      'special rate
'      lnMinRatex = IFNull(loRS("nSpclRate"), 0)
'   End If
'
'   'mac 2020.09.01
'   '  2H sales accepts 0 DP
'   p_nMinDownx = IFNull(loRS("nMinDownx"), 0#)
'
'   lsSQL = "a.sModelIDx = " & strParm(loRS("sModelIDx"))
'
'   lsSQL = "SELECT" & _
'               "  a.sModelIDx" & _
'               ", a.sModelNme" & _
'               ", b.nMinDownx" & _
'               ", c.nRebatesx" & _
'               ", c.nMiscChrg" & _
'               ", c.nEndMrtgg" & _
'               ", b.nSelPrice" & _
'               ", b.nLastPrce" & _
'               ", a.cHotItemx" & _
'            " FROM MC_Model a" & _
'               ", MC_Model_Price b" & _
'               ", MC_Category c" & _
'            " WHERE a.sModelIDx = b.sModelIDx" & _
'               " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
'               " AND b.sMCCatIDx = c.sMCCatIDx" & _
'               IIf(lsSQL <> Empty, " AND " & lsSQL, "")
'
'   Debug.Print lsSQL
'   If loRS.State = adStateOpen Then loRS.Close
'   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'
'   If loRS.EOF Then
'      GoTo endWithClear
'   ElseIf loRS.RecordCount > 1 Then
'      lsSelect = KwikBrowse(p_oAppDrivr, loRS _
'                           , "sModelIDx»sModelNme" _
'                           , "Code»Model Name")
'
'      If lsSelect <> "" Then
'         lasSelect = Split(lsSelect, "»")
'
'         loRS.MoveFirst
'         loRS.Find "sModelIDx = " & strParm(lasSelect(0)), 0, adSearchForward
'      End If
'   End If
'
'   p_sModelIDx = loRS("sModelIDx")
'   p_sModelNme = loRS("sModelNme")
'   p_nRebatesx = loRS("nRebatesx")
'   p_nMiscChrg = loRS("nMiscChrg")
'   p_nEndMrtgg = loRS("nEndMrtgg")
'   p_nSelPrice = loRS("nSelPrice")
'   p_nLastPrce = loRS("nLastPrce")
'
'   If lnMinRatex > 0# Then
'      'get the discount allowance
'      p_nSelPrice = IFNull(p_oAppDrivr.getConfiguration("2HRateAllw"), 0#)
'
'      'deduct the minumum discount rate
'      p_nSelPrice = loRS("nSelPrice") - (loRS("nSelPrice") * ((lnMinRatex - p_nSelPrice) / 100))
'
'      'round off to the nearest 50s as requested by mam loraine
'      p_nSelPrice = Round(p_nSelPrice / 50) * 50
'
'      'selling price will also be the last price
'      p_nLastPrce = loRS("nSelPrice") - (loRS("nSelPrice") * (lnMinRatex / 100))
'
'      'round off to the nearest 50s as requested by mam loraine
'      p_nLastPrce = Round(p_nLastPrce / 50) * 50
'   End If
'endProc:
'   ' XerSys - 2013-08-16
'   '  Reset the additional purchase for every model resets
'   p_nAddPurc = 0
'
'   p_cSoldStat = "1"
'   p_sSerialID = fsSerialID
'
'   Exit Sub
'endWithClear:
'   p_sModelIDx = ""
'   p_sModelNme = ""
'   p_nRebatesx = 0
'   p_nMiscChrg = 0
'   p_nEndMrtgg = 0
'
'   p_cSoldStat = "0"
'   p_sSerialID = ""
'
'   GoTo endProc
'errProc:
'   ShowError lsProcName & "( " & fsSerialID & " )"
'End Sub


