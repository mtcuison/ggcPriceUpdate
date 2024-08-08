VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCPurchasePriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Update MC Purchase Price History
'
' Copyright 2011 and Beyond
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
'  Kalyptus [ 12/21/2011 01:47 pm ]
'     Started coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMCCashPriceUpdate"

Private p_oAppDrivr As clsAppDriver
Private p_oMaster As Recordset

Private p_sSQLMastr As String


Private p_sBranchCd As String
Private p_sParentxx As String

Private pbInitTran As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Master = p_oMaster(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oMaster(Index) = Value
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
   
   p_sSQLMastr = "SELECT" & _
                     "  sModelIDx" & _
                     ", nPurPrice" & _
                     ", nSelPrice" & _
                     ", nLastPrce" & _
                     ", nDealrPrc" & _
                     ", nPromoPrc" & _
                     ", dPromoFrm" & _
                     ", dPromoTru" & _
                     ", sPromoIDx" & _
                     ", nMinDownx" & _
                     ", sMCCatIDx" & _
                     ", sApproved" & _
                     ", dPricexxx" & _
                     ", dPurchase" & _
                     ", dInsPrice" & _
                     ", cRecdStat" & _
                     ", sPrepared" & _
                     ", dPrepared" & _
                     ", sModified" & _
                     "  dModified" & _
                  " FROM MC_Model_Price"
   
   Set p_oMaster = New Recordset
   p_oMaster.Open AddCondition(p_sSQLMastr, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oMaster.ActiveConnection = Nothing
   p_oMaster.AddNew
   
   pbInitTran = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loHistory As Recordset
   Dim loRS As Recordset
   Dim loLowx As Recordset
   Dim loHigh As Recordset
   
   lsProcName = "SaveTransaction"
   On Error GoTo errProc
   
   lsSQL = AddCondition(p_sSQLMastr, "sModelIDx = " & strParm(p_oMaster("sModelIDx")))
   Set loHistory = New Recordset
   loHistory.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set loHistory.ActiveConnection = Nothing
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.BeginTrans
   
   'Check if model has a record in the MC_Model_Price table
   If loHistory.EOF = False Then
      'Model has record in MC_Model_Price
      'Check if purchase price is not null
      If IFNull(loHistory("nPurPrice")) <> 0 Then
         'Purchase Price is not null in the MC_Model_Price
         'Check if purchase price is different from the loaded record
         If loHistory("nPurPrice") <> p_oMaster("nPurPrice") Then
            'Purchase Price is Different from the price in the loaded record
                                    
            'Is purchase date of loaded record less than purchase date of the new record
            'It means a new Unit Price for this model
            If IFNull(loHistory("dPurchase"), "1900-01-01") < p_oMaster("dPurchase") Then
               'New Unit Price
               addHistory loHistory("sModelIDx"), loHistory("nPurPrice"), loHistory("dPurchase")
               
               'Update Unit Price of the Model
               lsSQL = "UPDATE MC_Model_Price SET " & _
                           " dPurchase = " & dateParm(IFNull(p_oMaster("dPurchase"), "1900-01-01")) & _
                           ", nPurPrice = " & p_oMaster("nPurPrice") & _
                      " WHERE sModelIDx = " & strParm(p_oMaster("sModelIDx"))
               p_oAppDrivr.Execute lsSQL, "MC_Model_Price_History"
            Else
               'Unit Price is an old Price
               
               Set loHigh = GetHistory(p_oMaster("sModelIDx"), p_oMaster("dPurchase"), True)
               Set loLowx = GetHistory(p_oMaster("sModelIDx"), p_oMaster("dPurchase"), False)
               
               'Check if there is a record for this item in the history with higher date
               If Not loHigh.EOF Then
                  'If same purchase price then we are dealing with the same record
                  If IFNull(loHigh("nPurPrice"), 0) = p_oMaster("nPurPrice") Then
                     'Is History higher than the new record
                     If (loHigh("dPurchase") > p_oMaster("dPurchase")) Then
                        lsSQL = "UPDATE MC_Model_Price_History SET" & _
                                    " dPurchase = " & dateParm(p_oMaster("dPurchase")) & _
                               " WHERE sModelIDx = " & strParm(p_oMaster("sModelIDx")) & _
                                 " AND dPurchase = " & dateParm(loHigh("dPurchase"))
                        p_oAppDrivr.Execute lsSQL, "MC_Model_Price_History"
                     End If
                  Else
                     If Not loLowx.EOF Then
                        If IFNull(loLowx("nPurPrice"), 0) <> p_oMaster("nPurPrice") Then
                           addHistory p_oMaster("sModelIDx"), p_oMaster("nPurPrice"), p_oMaster("dPurchase")
                        End If
                     Else
                        addHistory p_oMaster("sModelIDx"), p_oMaster("nPurPrice"), p_oMaster("dPurchase")
                     End If
                  End If
               ElseIf Not loLowx.EOF Then
                  'Unit Price History for this model can not be found
                  If IFNull(loLowx("nPurPrice"), 0) <> p_oMaster("nPurPrice") Then
                     addHistory p_oMaster("sModelIDx"), p_oMaster("nPurPrice"), p_oMaster("dPurchase")
                  End If
               Else
                  addHistory p_oMaster("sModelIDx"), p_oMaster("nPurPrice"), p_oMaster("dPurchase")
               End If
            End If
         Else
            '[If loHistory("nPurPrice") <> p_oMaster("nPurPrice")]
            '   -> Purchase price is equal with the loaded record
            'Check if purchase date from the loaded record is higher than the purchase date
            If IFNull(loHistory("dPurchase"), "1900-01-01") > p_oMaster("dPurchase") Then
               loHistory("nPurPrice") = p_oMaster("nPurPrice")
               loHistory("dPurchase") = p_oMaster("dPurchase")
               lsSQL = ADO2SQL(loHistory, _
                               "MC_Model_Price", _
                               "sModelIDx = " & strParm(p_oMaster("sModelIDx")), _
                               p_oAppDrivr.UserID, _
                               p_oAppDrivr.ServerDate)
               p_oAppDrivr.Execute lsSQL, "MC_Model_Price"
            End If
         End If
      Else
         '[If IFNull(loHistory("nPurPrice")) <> 0] -> Purchase Price is null in MC_Model_Price
         'Update MC_Model_Price
         loHistory("nPurPrice") = p_oMaster("nPurPrice")
         loHistory("dPurchase") = p_oMaster("dPurchase")
         lsSQL = ADO2SQL(loHistory, _
                         "MC_Model_Price", _
                         "sModelIDx = " & strParm(p_oMaster("sModelIDx")), _
                         p_oAppDrivr.UserID, _
                         p_oAppDrivr.ServerDate)
         p_oAppDrivr.Execute lsSQL, "MC_Model_Price"
      End If
   Else
      '[If loHistory.EOF] -> Model has no record in the MC_Model_Price
      'Create the model in the MC_Model_Price
      loHistory.AddNew
      loHistory("sModelIDx") = p_oMaster("sModelIDx")
      loHistory("nPurPrice") = p_oMaster("nPurPrice")
      loHistory("dPurchase") = p_oMaster("dPurchase")
            
      lsSQL = ADO2SQL(loHistory, "MC_Model_Price", , p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
            
      p_oAppDrivr.Execute lsSQL, "MC_Model_Price"
            
   End If
   
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.CommitTrans
   
   SaveTransaction = True
   
endProc:
   Exit Function
errProc:
   If Left(p_sParentxx, 3) <> "cls" Then p_oAppDrivr.RollbackTrans
   ShowError lsProcName & "( " & " )"
End Function

Private Function GetHistory(fsModelIDx As String, fdPurchase As Date, Optional fbHigher As Boolean = False)
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "GetHistory"
   On Error GoTo errProc
   
   lsSQL = "SELECT " & _
                   "  dPurchase" & _
                   ", nPurPrice" & _
          " FROM MC_Model_Price_History" & _
          " WHERE sModelIDx = " & strParm(p_oMaster("sModelIDx")) & _
            " AND dPurchase" & IIf(fbHigher, " > ", "<=") & dateParm(fdPurchase) & _
          " ORDER BY dPurchase " & IIf(fbHigher, " ASC ", " DESC ") & _
          " LIMIT 1"
   
   Set GetHistory = p_oAppDrivr.Connection.Execute(lsSQL, adCmdText)
    
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function addHistory(fsModelIDx As String, fnPurPrice As Currency, fdPurchase As Date) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "GetHistory"
   On Error GoTo errProc
   
   'Insert to history
   lsSQL = "INSERT INTO MC_Model_Price_History SET" & _
               "  sModelIDx = " & strParm(fsModelIDx) & _
               ", dPurchase = " & dateParm(fdPurchase) & _
               ", nPurPrice = " & fnPurPrice & _
               ", sModified = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
   p_oAppDrivr.Execute lsSQL, "MC_Model_Price_History"
      
   addHistory = True
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

