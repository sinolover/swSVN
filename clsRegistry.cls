VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function SetRegistryValue(hKey, KeyName, ValueName, Value, valueType, Flag) '404F50
  Dim var_2C As Me
  Dim var_ret_9 As Me
  loc_00404FCD: var_34 = ValueName
  loc_00404FD5: var_28 = Value
  loc_0040500A: var_ret_1 = KeyName
  loc_00405011: var_eax = RegCreateKeyEx(hKey, var_ret_1, 0, ebx, ebx, 63, var_4C, var_30, var_50)
  loc_0040502D: var_ret_2 = var_5C
  loc_00405035: var_50 = RegCreateKeyEx(hKey, var_ret_1, 0, ebx, ebx, 63, var_4C, var_30, var_50)
  loc_00405041: If var_50 <> 0 Then GoTo loc_0040527F
  loc_00405054: If VarType(var_28) > 17 Then GoTo loc_00405170
  loc_0040505A: If Not Asm.z_flag Then GoTo loc_0040512F
  loc_00405063: If VarType(var_28) < 2 Then GoTo loc_00405177
  loc_0040506C: If VarType(var_28) <= 3 Then GoTo loc_0040512F
  loc_00405075: If VarType(var_28) <> 8 Then GoTo loc_00405177
  loc_0040508A: var_40 = CStr(var_28)
  loc_00405095: eax = edx - 1
  loc_00405096: If edx - 1 = 0 Then GoTo loc_00405115
  loc_00405098: eax = edx - 1 - 1
  loc_0040509D: If var_40 = 0 Then GoTo loc_004050B5
  loc_004050AE: var_ret_4 = var_40
  loc_004050B3: GoTo loc_004050C9
  loc_004050C4: var_ret_5 = var_40
  loc_004050C9: 'Referenced from: 004050B3
  loc_004050D2: var_ret_6 = var_34
  loc_004050D9: var_eax = RegSetValueEx(var_30, var_ret_6, 0, 7, var_ret_5, Len(1), var_ret_4)
  loc_004050EE: var_ret_7 = var_5C
  loc_004050F8: var_ret_8 = var_60
  loc_00405104: var_50 = RegSetValueEx(var_30, var_ret_6, 0, 7, var_ret_5, Len(1), var_ret_4)
  loc_00405110: GoTo loc_00405261
  loc_00405115: 'Referenced from: 00405096
  loc_00405128: var_ret_9 = var_40
  loc_0040512D: GoTo loc_004050C9
  loc_0040512F: 'Referenced from: 0040505A
  loc_00405154: var_eax = RegSetValueEx(var_30, var_34, 0, 4, CLng(var_28), 4, 2)
  loc_00405169: var_ret_C = var_5C
  loc_0040516B: GoTo loc_00405255
  loc_00405170: 'Referenced from: 00405054
  loc_00405175: If var_ret_C = 8209 Then GoTo loc_0040518B
  loc_00405177: 
  loc_0040517B: var_eax = RegCloseKey(var_30)
  loc_00405186: GoTo loc_00405261
  loc_0040518B: 'Referenced from: 00405175
  loc_004051A1: var_2C = var_28
  loc_004051BB: var_ret_D = UBound(var_2C)
  loc_004051C4: var_ret_D = var_ret_D - LBound(1)
  loc_004051D6: var_3C = var_ret_D + 00000001h
  loc_004051DD: var_58 = var_2C
  loc_004051EA: If var_58 = 0 Then GoTo loc_0040520F
  loc_004051F0: If var_58 <> 1 Then GoTo loc_0040520F
  loc_004051FF: If Me.GetIDsOfNames < 0 Then GoTo loc_00405218
  loc_00405201: var_eax = Err.Raise
  loc_0040520D: GoTo loc_00405218
  loc_0040520F: 'Referenced from: 004051EA
  loc_0040520F: var_eax = Err.Raise
  loc_00405218: 'Referenced from: 0040520D
  loc_0040522D: var_ret_E = var_34
  loc_00405234: var_eax = RegSetValueEx(var_30, var_ret_E, 0, 3, Me.GetTypeInfoCount, var_3C, 0)
  loc_00405249: var_ret_F = var_5C
  loc_00405255: 'Referenced from: 0040516B
  loc_00405258: var_50 = RegSetValueEx(var_30, var_ret_E, 0, 3, Me.GetTypeInfoCount, var_3C, 0)
  loc_00405261: 'Referenced from: 00405110
  loc_00405265: var_eax = RegCloseKey(var_30)
  loc_00405277: setz al
  loc_0040527F: 'Referenced from: 00405041
  loc_00405284: GoTo loc_004052B0
  loc_004052AF: Exit Sub
  loc_004052B0: 'Referenced from: 00405284
End Function

Public Function GetRegistryValue(hKey, KeyName, ValueName, DefaultValue) '405310
  Dim var_34 As Me
  Dim var_98 As Me
  Dim var_4C As Me
  Dim var_1C As Me
  Dim var_50 As Me
  loc_00405394: var_38 = ValueName
  loc_004053A5: var_88 = IsMissing(DefaultValue)
  loc_004053D5: var_30 = IIf(IsMissing(DefaultValue), var_60, DefaultValue)
  loc_00405409: var_ret_1 = KeyName
  loc_00405410: var_eax = RegOpenKeyEx(hKey, var_ret_1, 0, 131097, var_20)
  loc_00405415: var_98 = RegOpenKeyEx(hKey, var_ret_1, 0, 131097, var_20)
  loc_00405429: var_ret_2 = var_50
  loc_0040543C: setnz dl
  loc_0040544D: If edx <> 0 Then GoTo loc_004058E3
  loc_0040546F: ReDim %r
  loc_00405480: var_4C = var_1C
  loc_0040548B: If var_4C = 0 Then GoTo loc_004054AC
  loc_00405491: If var_4C <> 1 Then GoTo loc_004054AC
  loc_0040549D: If Me.GetIDsOfNames < 0 Then GoTo loc_004054A8
  loc_0040549F: var_eax = Err.Raise
  loc_004054A8: 'Referenced from: 0040549D
  loc_004054AA: GoTo loc_004054B5
  loc_004054AC: 'Referenced from: 0040548B
  loc_004054AC: var_eax = Err.Raise
  loc_004054B5: 'Referenced from: 004054AA
  loc_004054CC: var_ret_3 = var_38
  loc_004054D3: var_eax = RegQueryValueEx(var_20, var_ret_3, 0, var_40, Me.GetTypeInfoCount, &H400, )
  loc_004054D8: var_98 = RegQueryValueEx(var_20, var_ret_3, 0, var_40, Me.GetTypeInfoCount, &H400, )
  loc_004054EC: var_ret_4 = var_50
  loc_00405511: If var_98 <> 234 Then GoTo loc_004055D4
  loc_0040551B: var_3C = var_3C - 00000001h
  loc_00405534: ReDim %r
  loc_0040554B: var_4C = var_1C
  loc_00405552: If var_4C = 0 Then GoTo loc_00405579
  loc_00405558: If var_4C <> 1 Then GoTo loc_00405579
  loc_00405564: If Me.GetIDsOfNames < 0 Then GoTo loc_0040556F
  loc_00405566: var_eax = Err.Raise
  loc_0040556F: 'Referenced from: 00405564
  loc_00405577: GoTo loc_00405582
  loc_00405579: 'Referenced from: 00405552
  loc_00405579: var_eax = Err.Raise
  loc_00405582: 'Referenced from: 00405577
  loc_00405599: var_ret_5 = var_38
  loc_004055A0: var_eax = RegQueryValueEx(var_20, var_ret_5, 0, var_40, Me.GetTypeInfoCount, var_3C - 00000001h, )
  loc_004055A5: var_98 = RegQueryValueEx(var_20, var_ret_5, 0, var_40, Me.GetTypeInfoCount, var_3C - 00000001h, )
  loc_004055B9: var_ret_6 = var_50
  loc_004055D2: GoTo loc_004055DA
  loc_004055D4: 'Referenced from: 00405511
  loc_004055DA: 'Referenced from: 004055D2
  loc_004055DD: eax = var_40 - 1
  loc_004055E1: If var_40 - 1 > 0 Then GoTo loc_0040587A
  loc_004055E7: GoTo loc_[eax*4+0040597Ch]
  loc_004055F6: var_4C = var_1C
  loc_004055FD: If var_4C = 0 Then GoTo loc_0040561E
  loc_00405603: If var_4C <> 1 Then GoTo loc_0040561E
  loc_0040560F: If Me.GetIDsOfNames < 0 Then GoTo loc_0040561A
  loc_00405611: var_eax = Err.Raise
  loc_0040561A: 'Referenced from: 0040560F
  loc_0040561C: GoTo loc_00405627
  loc_0040561E: 'Referenced from: 004055FD
  loc_0040561E: var_eax = Err.Raise
  loc_00405627: 'Referenced from: 0040561C
  loc_00405633: var_eax = CopyMemory(var_34, Me.GetTypeInfoCount, 4)
  loc_0040564B: var_78 = var_34
  loc_0040564E: GoTo loc_004056C7
  loc_00405658: var_4C = var_1C
  loc_0040565F: If var_4C = 0 Then GoTo loc_00405680
  loc_00405665: If var_4C <> 1 Then GoTo loc_00405680
  loc_00405671: If Me.GetIDsOfNames < 0 Then GoTo loc_0040567C
  loc_00405673: var_eax = Err.Raise
  loc_0040567C: 'Referenced from: 00405671
  loc_0040567E: GoTo loc_00405689
  loc_00405680: 'Referenced from: 0040565F
  loc_00405680: var_eax = Err.Raise
  loc_00405689: 'Referenced from: 0040567E
  loc_00405695: var_eax = CopyMemory(var_34, Me.GetTypeInfoCount, 4)
  loc_004056C7: 'Referenced from: 0040564E
  loc_004056D4: var_30 = var_98
  loc_004056DA: GoTo loc_004058D4
  loc_004056E2: var_3C = var_3C - 00000001h
  loc_004056F7: var_48 = Space$(var_3C - 00000001h)
  loc_00405705: var_4C = var_1C
  loc_0040570C: If var_4C = 0 Then GoTo loc_0040572D
  loc_00405712: If var_4C <> 1 Then GoTo loc_0040572D
  loc_0040571E: If Me.GetIDsOfNames < 0 Then GoTo loc_00405729
  loc_00405720: var_eax = Err.Raise
  loc_00405729: 'Referenced from: 0040571E
  loc_0040572B: GoTo loc_00405736
  loc_0040572D: 'Referenced from: 0040570C
  loc_0040572D: var_eax = Err.Raise
  loc_00405736: 'Referenced from: 0040572B
  loc_0040573C: var_3C = var_3C - 00000001h
  loc_00405754: var_eax = CopyMemory(var_48, Me.GetTypeInfoCount, var_3C - 00000001h)
  loc_00405767: var_ret_8 = var_50
  loc_00405784: If var_40 <> 2 Then GoTo loc_004057B8
  loc_00405794: Me.var_48 = Forms
  loc_004057AD: var_30 = var_50
  loc_004057B3: GoTo loc_004058D4
  loc_004057B8: 'Referenced from: 00405784
  loc_004057C2: var_78 = var_48
  loc_004057C5: GoTo loc_004058C8
  loc_004057CD: var_3C = var_3C - 00000002h
  loc_004057E2: var_48 = Space$(var_3C - 00000002h)
  loc_004057F0: var_eax = Me.GetIDsOfNames
  loc_004057F7: If var_4C = 0 Then GoTo loc_00405818
  loc_004057FD: If var_4C <> 1 Then GoTo loc_00405818
  loc_00405809: If Me.GetIDsOfNames < 0 Then GoTo loc_00405814
  loc_0040580B: var_eax = Err.Raise
  loc_00405814: 'Referenced from: 00405809
  loc_00405816: GoTo loc_00405821
  loc_00405818: 'Referenced from: 004057F7
  loc_00405818: var_eax = Err.Raise
  loc_00405821: 'Referenced from: 00405816
  loc_0040583F: var_eax = CopyMemory(var_48, Me.GetTypeInfoCount, var_3C - 00000002h)
  loc_00405852: var_ret_A = var_50
  loc_00405875: var_78 = var_48
  loc_00405878: GoTo loc_004058C8
  loc_0040587A: 'Referenced from: 004055E1
  loc_00405880: var_ret_B = UBound(var_1C)
  loc_0040588B: var_ret_B = var_ret_B + 00000001h
  loc_00405896: If var_3C = 0 Then GoTo loc_004058BB
  loc_00405898: var_3C = var_3C - 00000001h
  loc_004058B2: ReDim Preserve %r
  loc_004058BB: 'Referenced from: 00405896
  loc_004058C8: 'Referenced from: 004057C5
  loc_004058CE: var_30 =
  loc_004058D4: 'Referenced from: 004056DA
  loc_004058D8: var_eax = RegCloseKey(var_20)
  loc_004058E3: 'Referenced from: 0040544D
  loc_004058E8: GoTo loc_00405920
  loc_004058EE: If var_4 = 0 Then GoTo loc_004058F9
  loc_004058F9: 'Referenced from: 004058EE
  loc_0040591F: Exit Sub
  loc_00405920: 'Referenced from: 004058E8
  loc_00405941: Exit Sub
End Function

Public Function DeleteRegistryValueOrKey(hKey, RegKeyName, ValueName) '4059A0
  loc_004059F4: var_ret_1 = RegKeyName
  loc_004059FF: var_eax = RegOpenKeyEx(hKey, var_ret_1, 0, 63, var_28)
  loc_00405A0A: var_30 = RegOpenKeyEx(hKey, var_ret_1, 0, 63, var_28)
  loc_00405A1A: var_ret_2 = var_2C
  loc_00405A28: If var_30 <> 0 Then GoTo loc_00405A79
  loc_00405A34: var_ret_3 = ValueName
  loc_00405A3F: var_eax = RegDeleteValue(var_28, var_ret_3)
  loc_00405A44: var_30 = RegDeleteValue(var_28, var_ret_3)
  loc_00405A51: var_ret_4 = var_2C
  loc_00405A5F: If var_30 <> 0 Then GoTo loc_00405A79
  loc_00405A65: var_eax = RegCloseKey(var_28)
  loc_00405A70: If RegCloseKey(var_28) <> 0 Then GoTo loc_00405A79
  loc_00405A79: 'Referenced from: 00405A28
  loc_00405A7E: GoTo loc_00405A8A
  loc_00405A89: Exit Sub
  loc_00405A8A: 'Referenced from: 00405A7E
End Function

Public Sub Proc_1_3_405AC0
  loc_00405B2C: ExpandEnvironmentStrings(arg_C, vbNullString, var_ret_3 = #StkVar1%StkVar2)
  loc_00405B31: var_4C = var_ret_3
  loc_00405B45: var_ret_4 = var_24
  loc_00405B4F: var_ret_5 = var_28
  loc_00405B6A: var_4C = var_4C - 00000001h
  loc_00405B8E: var_18 = String$(var_4C, 0)
  loc_00405BB6: var_eax = ExpandEnvironmentStrings(arg_C, var_18, var_4C)
  loc_00405BC6: var_ret_8 = var_24
  loc_00405BD0: var_ret_9 = var_28
  loc_00405BEB: var_20 = var_18
  loc_00405BF6: GoTo loc_00405C24
  loc_00405BFC: If var_4 = 0 Then GoTo loc_00405C07
  loc_00405C07: 'Referenced from: 00405BFC
  loc_00405C23: Exit Sub
  loc_00405C24: 'Referenced from: 00405BF6
  loc_00405C2D: Exit Sub
End Sub

Public Sub Proc_1_4_405C60
  loc_00405C7D: VarPtr(esi.Release) = VarPtr(esi.Release) + 00000003h
  loc_00405C8A: CopyMemory(VarPtr(esi.Release), esi.%x1 = Invoke, 1)
  loc_00405CA7: VarPtr(esi.%x1 = Invoke) = VarPtr(esi.%x1 = Invoke) + 00000001h
  loc_00405CAE: VarPtr(esi.GetTypeInfoCount) = VarPtr(esi.GetTypeInfoCount) + 00000002h
  loc_00405CB5: CopyMemory(VarPtr(esi.GetTypeInfoCount), VarPtr(esi.%x1 = Invoke), 1)
  loc_00405CCC: VarPtr() = VarPtr() + 00000002h
  loc_00405CD3: VarPtr(esi.GetTypeInfoCount) = VarPtr(esi.GetTypeInfoCount) + 00000001h
  loc_00405CDA: var_eax = CopyMemory(VarPtr(esi.GetTypeInfoCount), VarPtr(), )
  loc_00405CE8: VarPtr() = VarPtr() + 00000003h
  loc_00405CF5: var_eax = CopyMemory(, , esi.GetTypeInfo)
End Sub
