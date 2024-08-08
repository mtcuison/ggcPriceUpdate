VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Update MC Price History
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
'  XerSys [ 03/30/2009 10:58 am ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMCSales"
Private Const pxeAddFreight As String = "AddFrgth"

Private p_oAppDrivr As clsAppDriver
Private p_oDetail As Recordset

Private p_sSQLMastr As String

Private p_sMCCatIDx As String
Private p_sMCCatNme As String
Private p_nMarkUpRt As Double
Private p_nMarkUpAm As Double
Private p_dPricexxx As Date
Private p_sApproved As String

Private p_sBranchCd As String
Private p_sParentxx As String
Private p_nAddFrgth As Double

Private p_cTranStat As xeTransactionStatus

Public Event DetailRetrieved(ByVal Index As Integer)
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
   If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
   p_oDetail.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, 3, 4, "nselprice", "nlastprce", "ndealrprc"
      If Not IsNumeric(Value) Then Exit Property
      
      p_oDetail(Index) = Value
   Case 5, "smccatnme"
      getDetail(
   End Select
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
   Case 2, "nmarkuprt"
      Master = p_nMarkUpRt
   Case 3, "nmarkupam"
      Master = p_nMarkUpAm
   Case 4, "dpricexxx"
      Master = p_dPricexxx
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "smccatidx"
   Case 1, "smccatnme"
      p_sMCCatNme = getMaster(Value)
   Case 2, "nmarkuprt"
      If Not IsNumeric(Value) Then Exit Property
      
      p_nMarkUpRt = Value
   Case 3, "nmarkupam"
      If Not IsNumeric(Value) Then Exit Property
      
      p_nMarkUpAm = Value
   Case 4, "dpricexxx"
      If Not IsDate(Value) Then Exit Property
   
      p_dPricexxx = Value
   End Select
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oDetail.RecordCount
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQLBrows As String
   
   lsProcName = "InitTransaction"
   'On Error Goto errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sSQLMastr = "SELECT" & _
                     "  a.sModelIDx" & _
                     ", b.sModelNme" & _
                     ", a.nSelPrice" & _
                     ", a.nLastPrce" & _
                     ", a.nDealrPrc" & _
                     ", c.sMCCatNme" & _
                     ", a.sApproved" & _
                     ", a.dPricexxx" & _
                     ", a.nMinDownx" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                     ", a.sMCCatIDx" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                  " ORDER BY b.sModelNme"
                  
   p_sMCCatIDx = ""
   p_sMCCatNme = ""
   p_nMarkUpRt = 0#
   p_nMarkUpAm = 0#
   p_dPricexxx = Date
   p_sApproved = ""
   
   pbInitTran = InitTransaction
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsAppvID As String, lsAppvName As String
   Dim lnAppvRights As Integer, lnAllowRights As Integer
   Dim lbGetApproval As Boolean
   Dim lnCtr As Integer, lnEntryNo As Integer
   
   lsProcName = "p_oBaseClas_WillSave"
   'On Error Goto errProc
   
   If pbShowMsg Then
      If MsgBox("Are you sure you want to Save this Transaction?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
   End If
   
   With p_oBaseClas
      If .isDetailModify = False Then
         SaveTransaction = True
         GoTo endProc
      End If
      
      If p_nReplAmtx > 0 And p_oSPSales.Master("nReplAmtx") <> p_nReplAmtx Then
         If GetSPReplacement() = False Then GoTo endProc
      End If
         
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "cGAwyStat") = pxeAdded Then
            lbGetApproval = True
            Exit For
         End If
      Next
      
      If lbGetApproval Then
         lnAllowRights = 255 'p_oAppDrivr.getConfiguration("MCGAway")
         If lnAllowRights And p_oAppDrivr.UserLevel = 0 Then
            lbGetApproval = MsgBox("User Doesn't Have the Right to Add New Giveaway!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes
            If GetApproval(p_oAppDrivr, _
                  lnAppvRights, _
                  lsAppvID, _
                  lsAppvName, _
                  p_oAppDrivr.MenuName) = False Then GoTo endProc
            
            If lnAllowRights And lnAppvRights = 0 Then
               MsgBox "Approving Officer Has no Right to Add New Giveaway!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
      
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      If .EditMode = xeModeUpdate Then
         lsSQL = "SELECT nEntryNox" & _
                  " FROM " & .DetailTable & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                  " ORDER BY nEntryNox DESC" & _
                  " LIMIT 1"
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
         If loRS.EOF Then
            lnEntryNo = 0
         Else
            lnEntryNo = loRS("nEntryNox")
         End If
      End If
      
      lnCtr = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nEntryNox") = 0 Then
            '.Detail(pnCtr, "nEntryNox") = lnEntryNo
            lnEntryNo = lnEntryNo + 1
         End If
         
         ' filter out remove and replace giveaway
         If .Detail(pnCtr, "cGAwyStat") = pxeAdded Or _
               .Detail(pnCtr, "cGAwyStat") = pxeOriginal Then
            If .Detail(pnCtr, "nGivenxxx") > 0 Then
               loSPTrans.Detail(lnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
               loSPTrans.Detail(lnCtr, "nQuantity") = .Detail(pnCtr, "nGivenxxx")
               loSPTrans.Detail(lnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
               loSPTrans.Detail(lnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
               loSPTrans.Detail(lnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
               loSPTrans.Detail(lnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
               
               lnCtr = lnCtr + 1
            End If
         End If
            
         If .EditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO MC_SO_GiveAways SET" & _
                        "  sTransNox = " & strParm(.Master("sTransNox")) & _
                        ", nEntryNox = " & pnCtr + 1 & _
                        ", sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                        ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                        ", nGivenxxx = " & strParm(.Detail(pnCtr, "nGivenxxx")) & _
                        ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
         Else
            lsSQL = "SELECT sPartsIDx" & _
                     " FROM MC_SO_GiveAways" & _
                     " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & pnCtr + 1

            Set loRS = New Recordset
            loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

            If .Detail(pnCtr, "nEntryNox") = 0 Then
               lsSQL = "INSERT INTO MC_SO_GiveAways SET" & _
                           "  sTransNox = " & strParm(.Master("sTransNox")) & _
                           ", nEntryNox = " & lnEntryNo & _
                           ", sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", nGivenxxx = " & strParm(.Detail(pnCtr, "nGivenxxx")) & _
                           ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
            Else
               lsSQL = "UPDATE MC_SO_GiveAways SET" & _
                           "  sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", nGivenxxx = " & strParm(.Detail(pnCtr, "nGivenxxx")) & _
                           ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                        " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & .Detail(pnCtr, "nEntryNox")
            End If
         End If

         If p_oAppDrivr.Execute(lsSQL, "MC_SO_GiveAways", p_sBranchCd) = 0 Then
            MsgBox "Unable to Update MC Sales GiveAways!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
      
      ' save SP Transaction if a valid record exists
      If lnCtr > 0 Then
         If loSPTrans.SalesGiveAway(.Master("sTransNox"), _
               .Master("dTransact"), _
               IIf(p_sParentxx = "clsMCSales", .EditMode, xeModeAddNew)) = False Then
            GoTo endProc
         End If
      End If
   End With

   With p_oSPSales
      If .Detail(0, "sPartsIDx") <> Empty Then
         .Master("dTransact") = p_oBaseClas.Master("dTransact")
         .Master("sClientID") = p_oBaseClas.Master("sClientID")
         .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
         .Master("sSourceCd") = pxeMCSales
         If .SaveTransaction() = False Then GoTo endProc

         If loSPTrans.InitTransaction() = False Then GoTo endProc
      ElseIf pbHasSPRep And .Detail(0, "sPartsIDx") = Empty Then
         If .DeleteTransaction = False Then GoTo endProc
      End If
   End With
      
   SaveTransaction = True
   
endProc:
   Set loSPTrans = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function RetrieveDetail() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "RetrieveDetail"
   'On Error Goto errProc

   If Not pbInitTran Then GoTo endProc

   lsSQL = p_sSQLMastr
   If p_sMCCatIDx <> Empty Then
      lsSQL = AddCondition(p_sSQLMastr, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
   End If
   
   Set p_oDetail = New Recordset
   p_oDetail.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
   OpenTransaction = Not p_oDetail.EOF
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function getCategory(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   'On Error Goto errProc
End Function

Private Function getDetail(ByVal lnRow As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "getDetail"
   'On Error Goto errProc
   
   With p_oBaseClas
      If lsValue <> Empty Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "b.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sBarrcode = " & strParm(Trim(lsValue))
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
            If lbSearch Then
               lsDetail = "b.sDescript LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sDescript = " & strParm(Trim(lsValue))
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      lsDetail = .getDetail(lnRow, lsDetail)
      If lsDetail = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
   
      lasDetail = Split(lsDetail, "»")
      .Detail(lnRow, "sPartsIDx") = lasDetail(0)
      .Detail(lnRow, "sBarrCode") = lasDetail(1)
      .Detail(lnRow, "sDescript") = lasDetail(2)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(3)
      .Detail(lnRow, "nQuantity") = 1
      .Detail(lnRow, "nGivenxxx") = 0
      .Detail(lnRow, "cGAwyStat") = pxeAdded
      .Detail(lnRow, "nSelPrice") = lasDetail(4)
      .Detail(lnRow, "nResvOrdr") = lasDetail(5)
      .Detail(lnRow, "nBackOrdr") = lasDetail(6)
      .Detail(lnRow, "nReorderx") = lasDetail(7)
      .Detail(lnRow, "nLedgerNo") = lasDetail(8)
   End With
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sPartsIDx") = ""
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nGivenxxx") = 0
      .Detail(lnRow, "cGAwyStat") = pxeOriginal
      .Detail(lnRow, "nSelPrice") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nReorderx") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   getDetail = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

