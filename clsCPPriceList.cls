VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     CP Price List
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
'  XerSys [ 12/07/2015 11:43 am ]
'     Start coding this object...
'  XerSys [ 03/15/2016 09:51 am ]
'     Start modifying this object which was only a copy of the mc pricelist
'  XerSys [ 08/01/2016 02:08 pm ]
'     Update Zero percent computation per branch type
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCPPiceList"

Private p_oAppDrivr As clsAppDriver
Private p_oCashPrice As Recordset
Private p_oInsPrice As Recordset

Private p_sModelCde As String
Private p_sModelNme As String
Private p_sModelIDx As String

Private p_nSelPrice As Double
Private p_nLastPrce As Double
Private p_nDiscount As Double
Private p_nMinDownx(4) As Double
Private p_nMaxDownx(4) As Double
Private p_nEndMrtgg As Double
Private p_nRebatesx As Double
Private p_nMiscChrg(4) As Double
Private p_nOthersxx As Double
Private p_nDownPaym(4) As Double
Private p_dTransact As Date

Private p_sSQLCash As String
Private p_sSQLInst As String
Private p_bExactNm As Boolean
Private p_bByCode As Boolean

Private p_bMFiesta As Boolean
Private p_xFactorRt(4) As Double

Private pbInitTran As Boolean

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

Property Get CPModelID() As String
   CPModelID = p_sModelCde
End Property

Property Let CPModelID(ByVal Value As String)
   p_bByCode = True
   Call getModel(Value)
End Property

Property Get CPModel() As String
   CPModel = p_sModelNme
End Property

Property Let CPModel(ByVal Value As String)
   p_bByCode = False
       Call getModel(Value)
End Property

Property Let ModelID(ByVal Value As String)
   Call getModelByID(Value)
End Property

Property Get ModelID() As String
   ModelID = p_sModelIDx
End Property

Property Get LastPrice() As Double
   LastPrice = p_nLastPrce
End Property

Property Get SelPrice() As Double
   SelPrice = p_nSelPrice
End Property

'mac 2022.02.16
'  accept change price, will be used for festival
Property Let SelPrice(ByVal fnValue As Double)
   p_nSelPrice = fnValue
End Property

Property Get Rebate() As Double
   Rebate = p_nRebatesx
End Property

Property Let Discount(ByVal fnValue As Double)
   p_nDiscount = fnValue
End Property

Property Get Discount() As Double
   Discount = p_nDiscount
End Property

Property Get MinimumDown(ByVal fnIndex As Integer) As Double
   Dim lnMinDown As Currency
   
   lnMinDown = p_nMinDownx(fnIndex) * (p_nSelPrice + p_nOthersxx - p_nDiscount)
   If lnMinDown - Int(lnMinDown) > 0 Then lnMinDown = Int(lnMinDown) + 1
   MinimumDown = lnMinDown + MiscCharge(fnIndex)
End Property

Property Get MaximumDown(ByVal fnIndex As Integer) As Double
   Dim lnMaxDown As Currency
   
   lnMaxDown = p_nMaxDownx(fnIndex) * (p_nSelPrice + p_nOthersxx - p_nDiscount)
   If lnMaxDown - Int(lnMaxDown) > 0 Then lnMaxDown = Int(lnMaxDown) + 1
   MaximumDown = lnMaxDown + MiscCharge(fnIndex)
End Property

Property Let DownPayment(ByVal fnIndex As Integer, ByVal fnValue As Double)
   Select Case fnValue
   Case MinimumDown(fnIndex) To MaximumDown(fnIndex)
      p_nDownPaym(fnIndex) = fnValue
   Case Else
      p_nDownPaym(fnIndex) = MinimumDown(fnIndex)
   End Select
End Property

Property Get DownPayment(ByVal fnIndex As Integer) As Double
   DownPayment = p_nDownPaym(fnIndex)
End Property

Property Get MiscCharge(ByVal fnIndex As Integer) As Double
   Dim lnMiscChrg As Currency
   
   lnMiscChrg = p_nMiscChrg(fnIndex) / 100 * (p_nSelPrice + p_nOthersxx - p_nDiscount)
   If lnMiscChrg - Int(lnMiscChrg) > 0 Then lnMiscChrg = Int(lnMiscChrg) + 1
   MiscCharge = lnMiscChrg
End Property

Property Get EndMortgage() As Double
   EndMortgage = p_nEndMrtgg
End Property

Property Get CashPriceCount() As Long
   If pbInitTran = False Then Exit Property
   
   CashPriceCount = p_oCashPrice.RecordCount
End Property

Property Get CashLatestDate() As Date
   Dim ldLatest As Date
   
   CashLatestDate = CDate("2016-01-01")
   With p_oCashPrice
      If .RecordCount > 1 Then
         .MoveFirst
         ldLatest = IFNull(.Fields("dPricexxx"), "2016-01-01")
         Do While .EOF = False
            If ldLatest < .Fields("dPricexxx") Then ldLatest = .Fields("dPricexxx")
            .MoveNext
         Loop
         CashLatestDate = ldLatest
      End If
   End With
End Property

Property Let OtherAmount(Value As Double)
   p_nOthersxx = Value
End Property

Property Let DateTransact(ByVal Value As Date)
   p_dTransact = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_sSQLCash = "SELECT b.sModelIDx" & _
                     ", b.sModelNme" & _
                     ", b.sModelCde" & _
                     ", c.sBrandNme" & _
                     ", a.nSelPrice" & _
                     ", a.nLastPrce" & _
                     ", a.dPricexxx" & _
                  " FROM CP_Model_Price a" & _
                     ", CP_Model b" & _
                     ", CP_Brand c" & _
                     ", CP_Inventory d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND b.sBrandIDx = c.sBrandIDx" & _
                     " AND a.sModelIDx = d.sModelIDx" & _
                     " AND d.sCategID1 IN('C001001','C001012','C0W1003','C0W1006','C0W1026') " & _
                     " AND d.cHsSerial = " & strParm(xeYes) & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND c.cRecdStat = " & strParm(xeRecStateActive) & _
                  " GROUP BY a.sModelIDx" & _
                  " ORDER BY c.sBrandNme, b.sModelNme, b.sModelCde"
   
   Set p_oCashPrice = New Recordset
   With p_oCashPrice
      .Open AddCondition(p_sSQLCash, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
   End With
   Debug.Print p_sSQLCash
   
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
   
   If p_sModelCde <> "" Then
   'she 2016-05-20 change code filter to ModelId may mga same modelCode kasi kaya iba bigay nya na price
      lsSQL = AddCondition(p_sSQLCash, "b.sModelIDx = " & strParm(p_sModelIDx))
'      lsSQL = AddCondition(p_sSQLCash, "b.sModelCde = " & strParm(p_sModelCde))
   Else
      lsSQL = p_sSQLCash
   End If
   
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

Function LoadZeroPercent() As Recordset
   Dim lsProcName As String
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lnSRP As Double
   Dim lsBrandIDx As String, lcShopType As String, lsAreaCode As String
   
   lsProcName = "LoadZeroPercent"
   On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   Call LoadCashPrice
   
   ' XerSys - 2016-08-01
   '  Retrieve type of branch
   lsSQL = "SELECT cShopType, sBrandIDx, sAreaCode" & _
            " FROM Branch_Others" & _
            " WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   If lors.EOF Then GoTo endProc
   lcShopType = lors("cShopType")
   lsBrandIDx = IFNull(lors("sBrandIDx"), "")
   lsAreaCode = lors("sAreaCode")
   
   lsSQL = "SELECT a.sBankIDxx" & _
               ", d.sBankName" & _
               ", d.sBankCode" & _
               ", c.sModelNme" & _
               ", c.sModelCde" & _
               ", a.n03MoTerm" & _
               ", a.n06MoTerm" & _
               ", a.n12MoTerm" & _
               ", a.n24MoTerm" & _
               ", a.n36MoTerm" & _
               ", a.nMin6Monx" & _
               ", a.nMin12Mon" & _
               ", a.nMin24Mon" & _
               ", b.cWith24Mo" & _
               ", b.cWith36Mo" & _
               ", b.n03MoTerm x03MoTerm" & _
               ", b.n06MoTerm x06MoTerm" & _
               ", b.n12MoTerm x12MoTerm" & _
               ", b.n24MoTerm x24MoTerm" & _
               ", b.n36MoTerm x36MoTerm"
   lsSQL = lsSQL & _
            " FROM CP_Card_Rate a" & _
                  " LEFT JOIN CP_Card_Rate_Model_Promo b" & _
                        " LEFT JOIN CP_Model c" & _
                           " ON b.sModelIDx = c.sModelIDx" & _
                     " ON a.sBankIDxx = b.sBankIDxx" & _
                        " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
                        " AND DATE(SYSDATE()) BETWEEN b.dPromoFrm" & _
                           " AND b.dPromoTru" & _
                        " AND ( b.cShopType = ''" & _
                           " OR ( b.cShopType = " & strParm(lcShopType) & _
                              " AND b.sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
                        " AND ( b.sAreaCode = ''" & _
                           " OR ( b.sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
                        " AND ( b.sBranchCd = ''" & _
                           " OR ( b.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
                        IIf(p_sModelIDx = "", "", " AND b.sModelIDx = " & strParm(p_sModelIDx)) & _
                  " LEFT JOIN Banks d" & _
                     " ON a.sBankIDxx = d.sBankIDxx" & _
            " WHERE a.cRecdStat = " & strParm(xeRecStateActive)
   lsSQL = lsSQL & " UNION SELECT b.sBankIDxx" & _
               ", d.sBankName" & _
               ", d.sBankCode" & _
               ", c.sModelNme" & _
               ", c.sModelCde" & _
               ", a.n03MoTerm" & _
               ", a.n06MoTerm" & _
               ", a.n12MoTerm" & _
               ", a.n24MoTerm" & _
               ", a.n36MoTerm" & _
               ", a.nMin6Monx" & _
               ", a.nMin12Mon" & _
               ", a.nMin24Mon" & _
               ", b.cWith24Mo" & _
               ", b.cWith36Mo" & _
               ", b.n03MoTerm x03MoTerm" & _
               ", b.n06MoTerm x06MoTerm" & _
               ", b.n12MoTerm x12MoTerm" & _
               ", b.n24MoTerm x24MoTerm" & _
               ", b.n36MoTerm x36MoTerm"
   lsSQL = lsSQL & _
            " FROM CP_Card_Rate_Model_Promo b" & _
                  " LEFT JOIN CP_Model c" & _
                     " ON b.sModelIDx = c.sModelIDx" & _
                  " LEFT JOIN Banks d" & _
                     " ON b.sBankIDxx = d.sBankIDxx" & _
                  " LEFT JOIN CP_Card_Rate a" & _
                     " ON a.sBankIDxx = b.sBankIDxx" & _
                        " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
            " WHERE b.cRecdStat = " & strParm(xeRecStateActive) & _
               " AND DATE(SYSDATE()) BETWEEN b.dPromoFrm" & _
                  " AND b.dPromoTru" & _
               " AND ( b.cShopType = ''" & _
                  " OR ( b.cShopType = " & strParm(lcShopType) & _
                     " AND b.sBrandIDx = " & strParm(lsBrandIDx) & " ) )" & _
               " AND ( b.sAreaCode = ''" & _
                  " OR ( b.sAreaCode = " & strParm(lsAreaCode) & " ) )" & _
               " AND ( b.sBranchCd = ''" & _
                  " OR ( b.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & " ) )" & _
               IIf(p_sModelIDx = "", "", " AND b.sModelIDx = " & strParm(p_sModelIDx)) & _
            " HAVING a.n03MoTerm IS NULL" & _
            " ORDER BY sModelNme"
   
   Debug.Print lsSQL
   Set p_oInsPrice = New Recordset
   With p_oInsPrice
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      If .EOF Then .AddNew
      
      Set lors = New Recordset
      Call createZeroPercent(lors)
      
      p_oCashPrice.MoveFirst
      Do Until p_oCashPrice.EOF
         Do Until .EOF()
            lors.AddNew
            lors("Model") = IIf(p_sModelNme = "", p_oCashPrice("sModelNme"), p_sModelNme)
            lors("Bank Name") = ""
            lors("Bank Code") = ""
            lors("3 Mos") = 0#
            lors("3 Gross") = 0#
            lors("6 Mos") = 0#
            lors("6 Gross") = 0#
            lors("12 Mos") = 0#
            lors("12 Gross") = 0#
            lors("24 Mos") = 0#
            lors("24 Gross") = 0#
            lors("36 Mos") = 0#
            lors("36 Gross") = 0#
            
            ' use the standard computation for each bank entries
            If IsNull(.Fields("sBankName")) Then
               lors("Bank Code") = "OTH"
               lors("Bank Name") = "Major Cards"
            Else
               lors("Bank Code") = .Fields("sBankCode")
               lors("Bank Name") = .Fields("sBankName")
            End If
            
            If IsNull(.Fields("sModelNme")) Then
               lors("3 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("n03MoTerm") / 100))
               lors("3 Mos") = lors("3 Gross") / 3
               
               If .Fields("nMin6Monx") <= p_oCashPrice("nSelPrice") Then
                  lors("6 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("n06MoTerm") / 100))
                  lors("6 Mos") = lors("6 Gross") / 6
               End If
               
               If (.Fields("nMin12Mon") > 0 And .Fields("nMin12Mon") <= p_oCashPrice("nSelPrice")) Or _
                     (.Fields("nMin12Mon") = 0 And .Fields("nMin6Monx") <= p_oCashPrice("nSelPrice")) Then
                  lors("12 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("n12MoTerm") / 100))
                  lors("12 Mos") = lors("12 Gross") / 12
               
                  If .Fields("n12MoTerm") <= .Fields("n24MoTerm") Then
                     lors("24 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("n24MoTerm") / 100))
                     lors("24 Mos") = lors("24 Gross") / 24
                  End If
               End If
               
               If (.Fields("nMin24Mon") > 0 And .Fields("nMin24Mon") <= p_oCashPrice("nSelPrice")) Or _
                     (.Fields("nMin24Mon") = 0 And .Fields("nMin12Mon") <= p_oCashPrice("nSelPrice")) Then
                  lors("24 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("n24MoTerm") / 100))
                  lors("24 Mos") = lors("24 Gross") / 24
                  
                  If .Fields("n24MoTerm") <= IFNull(.Fields("n36MoTerm"), 0) Then
                     lors("36 Gross") = p_oCashPrice("nSelPrice") * (1 + (IFNull(.Fields("n36MoTerm"), 0) / 100))
                     lors("36 Mos") = lors("36 Gross") / 36
                  End If
               End If
            Else ' Exceptional rate for the certain model
               lors("3 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("x03MoTerm") / 100))
               lors("3 Mos") = lors("3 Gross") / 3
'               If .Fields("nMin6Monx") <= p_oCashPrice("nSelPrice") Then
                  lors("6 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("x06MoTerm") / 100))
                  lors("6 Mos") = lors("6 Gross") / 6
'               End If
               
'               If (.Fields("nMin12Mon") > 0 And .Fields("nMin12Mon") <= p_oCashPrice("nSelPrice")) Or _
'                     (.Fields("nMin12Mon") = 0 And .Fields("nMin6Monx") <= p_oCashPrice("nSelPrice")) Then
                  lors("12 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("x12MoTerm") / 100))
                  lors("12 Mos") = lors("12 Gross") / 12
               
                  If .Fields("cWith24Mo") = xeYes Then
                     lors("24 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("x24MoTerm") / 100))
                     lors("24 Mos") = lors("24 Gross") / 24
                  End If
                  
                  If .Fields("cWith36Mo") = xeYes Then
                     lors("36 Gross") = p_oCashPrice("nSelPrice") * (1 + (.Fields("x36MoTerm") / 100))
                     lors("36 Mos") = lors("36 Gross") / 24
                  End If
'               End If
            End If
            .MoveNext
         Loop
         
         p_oCashPrice.MoveNext
      Loop
   End With
   p_bExactNm = False
   
   If lors.RecordCount > 0 Then lors.MoveFirst
   Set LoadZeroPercent = lors
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub createZeroPercent(lors As Recordset)
   Dim lsProcName As String
   
   lsProcName = "createZeroPercent"
   On Error GoTo errProc
   lors.Fields.Append "Model", adVarChar, 100, adFldIsNullable
   lors.Fields.Append "Bank Name", adVarChar, 64, adFldIsNullable
   lors.Fields.Append "Bank Code", adVarChar, 64, adFldIsNullable
   lors.Fields.Append "3 Mos", adDouble, 10, adFldIsNullable
   lors.Fields.Append "3 Gross", adDouble, 10, adFldIsNullable
   lors.Fields.Append "6 Mos", adDouble, 10, adFldIsNullable
   lors.Fields.Append "6 Gross", adDouble, 10, adFldIsNullable
   lors.Fields.Append "12 Mos", adDouble, 10, adFldIsNullable
   lors.Fields.Append "12 Gross", adDouble, 10, adFldIsNullable
   lors.Fields.Append "24 Mos", adDouble, 10, adFldIsNullable
   lors.Fields.Append "24 Gross", adDouble, 10, adFldIsNullable
   lors.Fields.Append "36 Mos", adDouble, 10, adFldIsNullable
   lors.Fields.Append "36 Gross", adDouble, 10, adFldIsNullable
   lors.Open
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub getModel(ByVal lsValue As String)
   Dim lors As Recordset
   Dim loRS1 As Recordset
   Dim loRS2 As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lsCondtion As String
   Dim lnCtr As Integer

   lsProcName = "getModel"
   On Error GoTo errProc
   
   If p_bByCode Then
      If lsValue = p_sModelCde Then GoTo endProc
      If p_bExactNm Then
         lsCondtion = "a.sModelIDx = " & strParm(Trim(lsValue))
      Else
         lsCondtion = "a.sModelIDx LIKE " & strParm(Trim(lsValue) & "%")
      End If
   Else
      If lsValue = p_sModelNme Then GoTo endProc
      If p_bExactNm Then
         lsCondtion = "f.sModelNme = " & strParm(Trim(lsValue))
      Else
         lsCondtion = "f.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
      End If
   End If
   
'   lsSQL = "SELECT a.sModelIDx" & _
'               ", a.sModelNme" & _
'               ", a.sModelCde" & _
'               ", b.nSelPrice" & _
'               ", b.nLastPrce" & _
'               ", b.nMinDownx" & _
'               ", b.nMaxDownx" & _
'               ", c.nRebatesx" & _
'               ", c.nEndMrtgg" & _
'               ", b.sMPCatIDx" & _
'            " FROM CP_Model a" & _
'               ", CP_Model_Price b" & _
'                  " LEFT JOIN CP_Category c" & _
'                     " ON b.sMPCatIDx = c.sMPCatIDx" & _
'            " WHERE a.sModelIDx = b.sModelIDx" & _
'               " AND b.nSelPrice > 0" & _
'               " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
'               IIf(lsSQL <> Empty, " AND " & lsSQL, "")
    
    lsSQL = "SELECT * FROM(SELECT" & _
               "  a.sModelIDx" & _
               ", a.nSelPrice" & _
               ", d.nSelPrice xSelPrice" & _
               ", a.nLastPrce" & _
               ", d.nLastPrce xLastPrce" & _
               ", a.nMinDownx" & _
               ", a.nMaxDownx" & _
               ", b.nRebatesx" & _
               ", b.nEndMrtgg" & _
               ", d.nMinDownx xMinDownx" & _
               ", d.nMaxDownx xMaxDownx" & _
               ", d.nMiscChrg xMiscChrg" & _
               ", d.sMPCatIDx xMPCatIDx" & _
               ", e.nRebatesx xRebatesx" & _
               ", e.nEndMrtgg xEndMrtgg" & _
               ", f.sModelNme" & _
               ", f.sModelCde" & _
               ", a.sMPCatIDx"
   lsSQL = lsSQL & _
            " FROM CP_Model_Price a" & _
               " LEFT JOIN CP_Model_Price_History d" & _
                  " LEFT JOIN CP_Category e" & _
                     " ON d.sMPCatIDx = e.sMPCatIDx" & _
                  " ON a.sModelIDx = d.sModelIDx" & _
                  " AND d.dPriceFrm <= " & dateParm(Format(p_oAppDrivr.ServerDate, "YYYY-MM-DD")) & _
                  " AND d.dPriceTru >= " & dateParm(Format(p_oAppDrivr.ServerDate, "YYYY-MM-DD")) & _
               " LEFT JOIN CP_Category b" & _
                  " ON a.sMPCatIDx = b.sMPCatIDx" & _
               ", CP_Model f" & _
            " WHERE a.sModelIDx = f.sModelIDx" & _
               IIf(lsCondtion <> Empty, " AND " & lsCondtion, "") & _
            " ORDER BY d.sMPCatIDx DESC)xxxSource" & _
            " GROUP BY sModelIDx"
            
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If lors.EOF Then
      GoTo endWithClear
   ElseIf lors.RecordCount >= 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, lors _
                           , "sModelCde»sModelNme»nSelPrice»sModelIDx" _
                           , "Code»Model Name»SelPrice»ID")

      If lsSelect <> "" Then
         lasSelect = Split(lsSelect, "»")
         
         lors.MoveFirst
         lors.Find "sModelIDx = " & strParm(lasSelect(0)), 0, adSearchForward
      Else
         GoTo endWithClear
      End If
   End If
   p_sModelIDx = lors("sModelIDx")
   p_sModelCde = lors("sModelCde")
   p_sModelNme = lors("sModelNme")
   p_nSelPrice = IFNull(lors("xSelPrice"), lors("nSelPrice"))
   p_nLastPrce = IFNull(lors("xLastPrce"), lors("nLastPrce"))

'   If IsNull(loRS("sMPCatIDx")) Then
'      lsSQL = "SELECT" & _
'               "  nMiscChrg" & _
'               ", nRebatesx" & _
'               ", nEndMrtgg" & _
'               ", nMinDownx" & _
'               ", nMaxDownx" & _
'            " FROM CP_Category" & _
'            " WHERE sMPCatIDx = 'C001000001'"
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'   End If

   p_nRebatesx = IFNull(lors("xRebatesx"), IFNull(lors("nRebatesx"), 0))
   p_nEndMrtgg = IFNull(lors("xEndMrtgg"), IFNull(lors("nEndMrtgg"), 0))
   p_nOthersxx = 0#
   p_nDiscount = 0#
   
   Debug.Print "SELECT" & _
                  "  nAcctThru" & _
                  ", nPrceFrom" & _
                  ", nPrceThru" & _
                  ", nMiscChrg" & _
               " FROM CP_Term_Category" & _
               " WHERE sMPCatIDx = " & strParm(IFNull(lors("xMPCatIDx"), IFNull(lors("sMPCatIDx"), ""))) & _
               " ORDER BY nAcctThru"
   
   Set loRS1 = New Recordset
   loRS1.Open "SELECT" & _
                  "  nAcctThru" & _
                  ", nPrceFrom" & _
                  ", nPrceThru" & _
                  ", nMiscChrg" & _
               " FROM CP_Term_Category" & _
               " WHERE sMPCatIDx = " & strParm(IFNull(lors("xMPCatIDx"), IFNull(lors("sMPCatIDx"), ""))) & _
               " ORDER BY nAcctThru" _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   p_nMiscChrg(0) = 0#
   p_nMiscChrg(1) = 0#
   p_nMiscChrg(2) = 0#
   p_nMiscChrg(3) = 0#
   p_nMiscChrg(4) = 0#
   
   p_nMinDownx(0) = 0#
   p_nMinDownx(1) = 0#
   p_nMinDownx(2) = 0#
   p_nMinDownx(3) = 0#
   p_nMinDownx(4) = 0#
   
   p_nMaxDownx(0) = 0#
   p_nMaxDownx(1) = 0#
   p_nMaxDownx(2) = 0#
   p_nMaxDownx(3) = 0#
   p_nMaxDownx(4) = 0#
   
   If p_oAppDrivr.BranchCode = "C0M2" Then
      GoTo setDefaults
   Else
setDefaults:
      For lnCtr = 0 To loRS1.RecordCount - 1
         Select Case CInt(loRS1("nAcctThru"))
         Case 3
            p_nMiscChrg(0) = loRS1("nMiscChrg")
            p_nMinDownx(0) = lors("nMinDownx")
            p_nMaxDownx(0) = lors("nMaxDownx")
         Case 6
            p_nMiscChrg(1) = loRS1("nMiscChrg")
            p_nMinDownx(1) = lors("nMinDownx")
            p_nMaxDownx(1) = lors("nMaxDownx")
         Case 9
            p_nMiscChrg(2) = loRS1("nMiscChrg")
            p_nMinDownx(2) = lors("nMinDownx")
            p_nMaxDownx(2) = lors("nMaxDownx")
         Case 12
            p_nMiscChrg(3) = loRS1("nMiscChrg")
            p_nMinDownx(3) = lors("nMinDownx")
            p_nMaxDownx(3) = lors("nMaxDownx")
         Case 24
            p_nMiscChrg(4) = loRS1("nMiscChrg")
            p_nMinDownx(4) = lors("nMinDownx")
            p_nMaxDownx(4) = lors("nMaxDownx")
         End Select
         loRS1.MoveNext
      Next
   End If
   
   
   
   

endProc:
   Exit Sub
endWithClear:
   p_sModelCde = ""
   p_sModelNme = ""
   p_nEndMrtgg = 0#
   p_nRebatesx = 0#
   
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue & " )"
End Sub

Private Sub getModelByID(ByVal lsValue As String)
   Dim lors As Recordset
   Dim loRS1 As Recordset
   Dim lsProcName As String
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   Dim lnMinDownx As Double
   Dim lnMaxDownx As Double
   Dim lnMiscChrg As Double

   lsProcName = "getModelByID"
   On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  a.nSelPrice" & _
               ", d.nSelPrice xSelPrice" & _
               ", a.nLastPrce" & _
               ", d.nLastPrce xLastPrce" & _
               ", a.nMinDownx" & _
               ", a.nMaxDownx" & _
               ", b.nRebatesx" & _
               ", b.nEndMrtgg" & _
               ", d.nMinDownx xMinDownx" & _
               ", d.nMaxDownx xMaxDownx" & _
               ", d.nMiscChrg xMiscChrg" & _
               ", d.sMPCatIDx xMPCatIDx" & _
               ", e.nRebatesx xRebatesx" & _
               ", e.nEndMrtgg xEndMrtgg" & _
               ", f.sModelNme" & _
               ", f.sModelCde" & _
               ", f.sModelIDx" & _
               ", a.sMPCatIDx"
   lsSQL = lsSQL & _
            " FROM CP_Model_Price a" & _
               " LEFT JOIN CP_Model_Price_History d" & _
                  " LEFT JOIN CP_Category e" & _
                     " ON d.sMPCatIDx = e.sMPCatIDx" & _
                  " ON a.sModelIDx = d.sModelIDx" & _
                  " AND d.dPriceFrm <= " & dateParm(Format(p_dTransact, "YYYY-MM-DD")) & _
                  " AND d.dPriceTru >= " & dateParm(Format(p_dTransact, "YYYY-MM-DD")) & _
               ", CP_Category b" & _
               ", CP_Model f" & _
            " WHERE a.sModelIDx = f.sModelIDx" & _
               " AND a.sMPCatIDx = b.sMPCatIDx" & _
               " AND a.sModelIDx = " & strParm(lsValue) & _
            " ORDER BY d.sMPCatIDx DESC" & _
            " LIMIT 1"

   Set lors = New Recordset
   Debug.Print lsSQL

   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If lors.EOF Then
      GoTo endWithClear
   ElseIf lors.RecordCount > 1 Then
      lsSelect = KwikBrowse(p_oAppDrivr, lors _
                           , "sModelIDx»sModelCde»sModelNme" _
                           , "ID»Code»Model Name")

      If lsSelect <> "" Then
         lasSelect = Split(lsSelect, "»")
         
         lors.MoveFirst
         lors.Find "sModelIDx = " & strParm(lasSelect(0)), 0, adSearchForward
      Else
         GoTo endWithClear
      End If
   End If
   
   p_sModelIDx = lors("sModelIDx")
   p_sModelCde = lors("sModelCde")
   p_sModelNme = lors("sModelNme")
   
   If InStr(1, "c0m2»c0a9", LCase(p_oAppDrivr.BranchCode), vbTextCompare) > 0 Then
      p_nLastPrce = p_nSelPrice
   Else
      p_nSelPrice = IFNull(lors("xSelPrice"), lors("nSelPrice"))
      p_nLastPrce = IFNull(lors("xLastPrce"), lors("nLastPrce"))
   End If
   
'   If IsNull(loRS("sMPCatIDx")) Then
'      lsSQL = "SELECT" & _
'               "  nMiscChrg" & _
'               ", nRebatesx" & _
'               ", nEndMrtgg" & _
'               ", nMinDownx" & _
'               ", nMaxDownx" & _
'            " FROM CP_Category" & _
'            " WHERE sMPCatIDx = 'C001000001'"
'
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'   End If
   
   p_nRebatesx = IFNull(lors("xRebatesx"), lors("nRebatesx"))
   p_nEndMrtgg = IFNull(lors("xEndMrtgg"), lors("nEndMrtgg"))
   p_nOthersxx = 0#
   p_nDiscount = 0#
      
   lsSQL = "SELECT" & _
                  "  nAcctThru" & _
                  ", nPrceFrom" & _
                  ", nPrceThru" & _
                  ", nMiscChrg" & _
               " FROM CP_Term_Category" & _
               " WHERE sMPCatIDx = " & strParm(IFNull(lors("xMPCatIDx"), lors("sMPCatIDx"))) & _
               " ORDER BY nAcctThru"
      
   Set loRS1 = New Recordset
   Debug.Print lsSQL
   loRS1.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   p_nMiscChrg(0) = 0#
   p_nMiscChrg(1) = 0#
   p_nMiscChrg(2) = 0#
   p_nMiscChrg(3) = 0#
   p_nMiscChrg(4) = 0#
   
   p_nMinDownx(0) = 0#
   p_nMinDownx(1) = 0#
   p_nMinDownx(2) = 0#
   p_nMinDownx(3) = 0#
   p_nMinDownx(4) = 0#
   
   p_nMaxDownx(0) = 0#
   p_nMaxDownx(1) = 0#
   p_nMaxDownx(2) = 0#
   p_nMaxDownx(3) = 0#
   p_nMaxDownx(4) = 0#
   
   p_bMFiesta = False
   
   If LCase(p_oAppDrivr.BranchCode) = "c0m2" Then
      p_bMFiesta = getMFRate(lsValue, lnMinDownx, lnMaxDownx, lnMiscChrg)
      
      If Not p_bMFiesta Then GoTo setDefaults
      
      p_nMiscChrg(0) = lnMiscChrg
      p_nMiscChrg(1) = lnMiscChrg
      p_nMiscChrg(2) = lnMiscChrg
      p_nMiscChrg(3) = lnMiscChrg
      p_nMiscChrg(4) = lnMiscChrg
      
      p_nMinDownx(0) = lnMinDownx
      p_nMinDownx(1) = lnMinDownx
      p_nMinDownx(2) = lnMinDownx
      p_nMinDownx(3) = lnMinDownx
      p_nMinDownx(4) = lnMinDownx
      
      p_nMaxDownx(0) = lnMaxDownx
      p_nMaxDownx(1) = lnMaxDownx
      p_nMaxDownx(2) = lnMaxDownx
      p_nMaxDownx(3) = lnMaxDownx
      p_nMaxDownx(4) = lnMaxDownx
   Else
setDefaults:
      For lnCtr = 0 To loRS1.RecordCount - 1
         Select Case CInt(loRS1("nAcctThru"))
         Case 3
            p_nMiscChrg(0) = loRS1("nMiscChrg")
            p_nMinDownx(0) = lors("nMinDownx")
            p_nMaxDownx(0) = lors("nMaxDownx")
         Case 6
            p_nMiscChrg(1) = loRS1("nMiscChrg")
            p_nMinDownx(1) = lors("nMinDownx")
            p_nMaxDownx(1) = lors("nMaxDownx")
         Case 9
            p_nMiscChrg(2) = loRS1("nMiscChrg")
            p_nMinDownx(2) = lors("nMinDownx")
            p_nMaxDownx(2) = lors("nMaxDownx")
         Case 12
            p_nMiscChrg(3) = loRS1("nMiscChrg")
            p_nMinDownx(3) = lors("nMinDownx")
            p_nMaxDownx(3) = lors("nMaxDownx")
         Case 24
            p_nMiscChrg(4) = loRS1("nMiscChrg")
            p_nMinDownx(4) = lors("nMinDownx")
            p_nMaxDownx(4) = lors("nMaxDownx")
         End Select
         loRS1.MoveNext
      Next
   End If
   
endProc:
   Exit Sub
endWithClear:
   p_sModelCde = ""
   p_sModelNme = ""
   p_nEndMrtgg = 0
   p_nRebatesx = 0
   
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue & " )"
End Sub

Function getMonthly(ByRef DownPayment As Double, _
                     Term As Integer, _
                     ByRef SelPrice As Currency, _
                     ByRef PrceFrom As Currency, _
                     ByRef PrceThru As Currency) As Double
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnService As Currency
   Dim lnDownPmt As Currency
   Dim lnSelPrce As Currency
   Dim lnMPAmtxx As Currency
   Dim lsMPCatID As String
   
   lsProcName = "getMonthly"
'   On Error GoTo errProc
   
   If p_sModelIDx = "" Then
      MsgBox "No Model has been Selected!" & vbCrLf & _
               "Please select a model first then try again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
               "  a.nSelPrice" & _
               ", d.nSelPrice xSelPrice" & _
               ", b.nRebatesx" & _
               ", b.nEndMrtgg" & _
               ", e.nRebatesx xRebatesx" & _
               ", e.nEndMrtgg xEndMrtgg" & _
               ", f.nAcctThru xAcctThru" & _
               ", f.nFactorRt xFactorRt" & _
               ", f.nPrceFrom xPrceFrom" & _
               ", f.nPrceThru xPrceThru" & _
               ", f.nMiscChrg xMiscChrg" & _
               ", f.sMPCatIDx xMPCatIDx" & _
               ", a.sMPCatIDx" & _
               ", g.nPrceFrom" & _
               ", g.nPrceThru" & _
               ", g.nMiscChrg" & _
               ", g.nFactorRt" & _
               ", g.nAcctThru" & _
               ", a.dPricexxx" & _
               ", a.dPriceTru" & _
               ", b.cZeroIntx"
               
   lsSQL = lsSQL & _
            " FROM CP_Model_Price a" & _
               " LEFT JOIN CP_Model_Price_History d" & _
                  " LEFT JOIN CP_Category e" & _
                     " ON d.sMPCatIDx = e.sMPCatIDx" & _
                  " LEFT JOIN CP_Term_Category f" & _
                     " ON d.sMPCatIDx = f.sMPCatIDx" & _
                     " AND f.nAcctThru = " & Term & _
                  " ON a.sModelIDx = d.sModelIDx" & _
                  " AND d.dPriceFrm <= " & dateParm(Format(p_dTransact, "YYYY-MM-DD")) & _
                  " AND d.dPriceTru >= " & dateParm(Format(p_dTransact, "YYYY-MM-DD")) & _
               ", CP_Category b" & _
                  " LEFT JOIN CP_Term_Category g" & _
                     " ON b.sMPCatIDx = g.sMPCatIDx" & _
                     " AND g.nAcctThru = " & Term & _
            " WHERE a.sMPCatIDx = b.sMPCatIDx" & _
               " AND a.sModelIDx = " & strParm(p_sModelIDx)
               
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
'   lsSQL = ""
'   If loRS.EOF Then
'      'Used regular rate instead
'      lsSQL = "SELECT" & _
'                  "  a.nSelPrice" & _
'                  ", b.nRebatesx" & _
'                  ", b.nEndMrtgg" & _
'                  ", c.nMiscChrg" & _
'                  ", c.nAcctThru" & _
'                  ", c.nFactorRt" & _
'                  ", c.nPrceFrom" & _
'                  ", c.nPrceThru" & _
'               " FROM CP_Model_Price a" & _
'                  ", CP_Category b" & _
'                  ", CP_Term_Category c" & _
'               " WHERE a.sMPCatIDx = b.sMPCatIDx" & _
'                  " AND a.sMPCatIDx = c.sMPCatIDx" & _
'                  " AND a.sModelIDx = " & strParm(p_sModelIDx) & _
'                  " AND c.nAcctThru = " & Term & _
'                  " AND a.sMPCatIDx = 'C001000001'"
'   Else
'      If IsNull(loRS("xAcctThru")) And IsNull(loRS("NAcctThru")) Then
'         lsSQL = "SELECT" & _
'                  "  a.nRebatesx" & _
'                  ", a.nEndMrtgg" & _
'                  ", b.nMiscChrg" & _
'                  ", b.nAcctThru" & _
'                  ", b.nFactorRt" & _
'                  ", b.nPrceFrom" & _
'                  ", b.nPrceThru" & _
'                  ", NULL xSelPrice" & _
'                  ", NULL xPrceFrom" & _
'                  ", NULL xPrceThru" & _
'                  ", NULL xMiscChrg" & _
'                  ", NULL xFactorRt" & _
'                  ", NULL xRebatesx" & _
'                  ", NULL xEndMrtgg" & _
'                  ", a.sMPCatIDx" & _
'               " FROM CP_Category a" & _
'                  ", CP_Term_Category b" & _
'               " WHERE a.sMPCatIDx = b.sMPCatIDx" & _
'                  " AND b.nAcctThru = " & Term & _
'                  " AND b.sMPCatIDx = 'C001000001'"
'      End If
'   End If
   
'   If lsSQL <> "" Then
'      MsgBox "No Record found to the given criteria..." & vbCrLf & _
'               "Regular rate set automatically...", vbInformation, "INFO"
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'
'      Dim lors1 As Recordset
'      Set lors1 = New Recordset
'      lors1.Open "SELECT" & _
'                     " nSelPrice" & _
'                  " FROM CP_Model_Price" & _
'                  " WHERE sModelIDx = " & strParm(p_sModelIDx) _
'      , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'      lnSelPrce = lors1("nSelPrice")
'   Else
      'mac 2023.02.16
      If InStr(1, "c0m2»c0a9", LCase(p_oAppDrivr.BranchCode), vbTextCompare) > 0 Then
         lnSelPrce = p_nSelPrice
      Else
         If lors.EOF Then
            lnSelPrce = 0#
         Else
            lnSelPrce = lors("nSelPrice")
         End If
      End If

''      If loRS.EOF Then
''         lnSelPrce = 0#
''      Else
''         lnSelPrce = loRS("nSelPrice")
''      End If
'   End If
   
   If lors.EOF Then GoTo endProc
   
   If Not IsNull(lors("xAcctThru")) Then
     If CInt(lors("xAcctThru")) <> Term Then
        getMonthly = 0#
        Exit Function
     End If
   Else
     If Not IsNull(lors("nAcctThru")) Then
        If CInt(lors("nAcctThru")) <> Term Then
            getMonthly = 0#
            Exit Function
        Else
            If CDate(lors("dPricexxx")) <= Format(p_oAppDrivr.ServerDate, "MM/DD/YYYY") And _
               CDate(IFNull(lors("dPriceTru"), p_oAppDrivr.ServerDate)) >= Format(p_oAppDrivr.ServerDate, "MM/DD/YYYY") Then
            Else
               getMonthly = 0#
               Exit Function
            End If
        End If
    End If
   End If
   
   If p_nOthersxx > (IFNull(lors("xSelPrice"), lnSelPrce) * 0.1) Then
      MsgBox "Unable to add other item!!!" & vbCrLf & _
               "Please verify your entry then try again!!!", vbCritical, "WARNING"
      p_nOthersxx = 0#
      GoTo endProc
   End If
   
'   If loRS("cZeroIntx") = 0 Then
    If IsNull(IFNull(lors("xPrceFrom"), lors("nPrceFrom"))) Then
       getMonthly = 0#
       GoTo endProc
    End If
'   End If
   
   PrceFrom = IFNull(lors("xPrceFrom"), IFNull(lors("nPrceFrom"), 0))
   PrceThru = IFNull(lors("xPrceThru"), IFNull(lors("nPrceThru"), 0))
   lsMPCatID = IFNull(lors("xMPCatIDx"), lors("sMPCatIDx"))
   
   'mac 2023.02.16
   If InStr(1, "c0m2»c0a9", LCase(p_oAppDrivr.BranchCode), vbTextCompare) > 0 Then
      SelPrice = p_nSelPrice
   Else
      Select Case IFNull(lors("xSelPrice"), lnSelPrce)
      Case Is >= lors("nPrceFrom")
         SelPrice = IFNull(lors("xSelPrice"), lnSelPrce)
      Case Else
         If lors("nPrceFrom") > 0# Then
            SelPrice = 0#
            GoTo endProc
         Else
            SelPrice = IFNull(lors("xSelPrice"), lnSelPrce)
         End If
      End Select
   End If
   
'   Debug.Print (SelPrice - DownPayment - MiscCharge) * loRS("nFactorRt") / Term
'   p_nMiscChrg = IFNull(loRS("xMiscChrg"), loRS("nMiscChrg"))
   
   If IFNull(lors("xFactorRt"), lors("nFactorRt")) > 0 Then
      Select Case Term
      Case 3
         If p_bMFiesta Then
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(0))) _
                   * p_xFactorRt(0) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         Else
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(0))) _
                   * IFNull(lors("xFactorRt"), lors("nFactorRt")) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         End If
      Case 6
         If p_bMFiesta Then
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(1))) _
                      * p_xFactorRt(1) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         Else
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(1))) _
                      * IFNull(lors("xFactorRt"), lors("nFactorRt")) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         End If
      Case 9
         If p_bMFiesta Then
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(2))) _
                      * p_xFactorRt(2) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         Else
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(2))) _
                      * IFNull(lors("xFactorRt"), lors("nFactorRt")) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         End If
      Case 12
         If p_bMFiesta Then
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(3))) _
                      * p_xFactorRt(3) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         Else
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(3))) _
                      * IFNull(lors("xFactorRt"), lors("nFactorRt")) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         End If
      Case 24
         If p_bMFiesta Then
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(4))) _
                      * p_xFactorRt(4) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         Else
            lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(4))) _
                      * IFNull(lors("xFactorRt"), lors("nFactorRt")) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
         End If
      End Select
      If lnMPAmtxx - Int(lnMPAmtxx) > 0 Then lnMPAmtxx = Int(lnMPAmtxx) + 1
   Else
      Select Case Term
      Case 3
         lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(0))) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
      Case 6
         lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(1))) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
      Case 9
         lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(2))) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
      Case 12
         lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(3))) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
      Case 24
         lnMPAmtxx = (((SelPrice + p_nOthersxx - p_nDiscount) - (DownPayment - MiscCharge(4))) / Term) + IFNull(lors("xRebatesx"), lors("nRebatesx")) + (IFNull(lors("xEndMrtgg"), lors("nEndMrtgg")) / Term)
      End Select
      If lnMPAmtxx - Int(lnMPAmtxx) > 0 Then lnMPAmtxx = Int(lnMPAmtxx) + 1
   End If
   
   Set lors = New Recordset
   lors.Open "SELECT" & _
                  " nMinMonMP" & _
               " FROM CP_Category" & _
               " WHERE sMPCatIDx = " & strParm(lsMPCatID) _
   , p_oAppDrivr.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lors("nMinMonMP") > 0# Then
      If lnMPAmtxx < CDbl(lors("nMinMonMP")) Then
         Select Case Term
         Case 3
            If p_nDownPaym(0) > MinimumDown(0) Then
            Else
               MsgBox "Monthly Payment is less than the minimum requirement!!!", vbCritical, "WARNING"
               lnMPAmtxx = 0#
            End If
         Case 6
            If p_nDownPaym(1) > MinimumDown(1) Then
            Else
               MsgBox "Monthly Payment is less than the minimum requirement!!!", vbCritical, "WARNING"
               lnMPAmtxx = 0#
            End If
         Case 9
            If p_nDownPaym(2) > MinimumDown(2) Then
            Else
               MsgBox "Monthly Payment is less than the minimum requirement!!!", vbCritical, "WARNING"
               lnMPAmtxx = 0#
            End If
         Case 12
            If p_nDownPaym(3) > MinimumDown(3) Then
            Else
               MsgBox "Monthly Payment is less than the minimum requirement!!!", vbCritical, "WARNING"
               lnMPAmtxx = 0#
            End If
         Case 24
            If p_nDownPaym(4) > MinimumDown(4) Then
            Else
               MsgBox "Monthly Payment is less than the minimum requirement!!!", vbCritical, "WARNING"
               lnMPAmtxx = 0#
            End If
         End Select
      End If
   End If
   
   getMonthly = lnMPAmtxx
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & DownPayment _
                        & ", " & Term & " )"
End Function

Private Function getMFRate(ByVal lsModelIDx As String, _
                        ByRef lnMinDownx As Double, _
                        ByRef lnMaxDownx As Double, _
                        ByRef lnMiscChrg As Double) As Boolean
   Dim lsMFPeriod As String
   Dim lsSQL As String
   Dim lors As Recordset
      
   p_xFactorRt(0) = 0#
   p_xFactorRt(1) = 0#
   p_xFactorRt(2) = 0#
   p_xFactorRt(3) = 0#
   p_xFactorRt(4) = 0#
      
   getMFRate = False
   lsSQL = "SELECT sFiestaCD" & _
          " FROM Mobile_Fiesta_Period" & _
          " WHERE " & dateParm(p_dTransact) & " BETWEEN dDateFrom AND dDateThru" & _
            " AND cRecdStat = '1'"
   Set lors = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If lors.EOF Then Exit Function
   
   lsMFPeriod = lors("sFiestaCD")
         
   lsSQL = "SELECT" & _
               "  nMinDownx" & _
               ", nMaxDownx" & _
               ", nMiscChrg" & _
               ", n03MoTerm" & _
               ", n06MoTerm" & _
               ", n12MoTerm" & _
               ", n24MoTerm" & _
               ", n36MoTerm" & _
               ", nUnitprce" & _
            " FROM Mobile_Fiesta_Model" & _
            " WHERE sModelIDx = " & strParm(lsModelIDx) & _
               " AND sFiestaCD = " & strParm(lsMFPeriod) & _
            " HAVING nUnitprce IS NOT NULL"
            
   Debug.Print lsSQL
   Set lors = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If lors.EOF Then Exit Function
   
   lnMinDownx = lors("nMinDownx")
   lnMaxDownx = lors("nMaxDownx")
   lnMiscChrg = lors("nMiscChrg")
   
   p_xFactorRt(0) = lors("n03MoTerm")
   p_xFactorRt(1) = lors("n06MoTerm")
   p_xFactorRt(2) = lors("n12MoTerm")
   p_xFactorRt(3) = lors("n24MoTerm")
   p_xFactorRt(4) = lors("n36MoTerm")
   
   getMFRate = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub
