VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const SecurityAnonymous As Long = 1
Private Declare Function apiIsUserAnAdmin Lib "shell32" Alias "IsUserAnAdmin" () As Long
'Private Declare Function apiCreatePrivateObjectSecurity Lib "ADVAPI32" Alias "CreatePrivateObjectSecurity" (ByRef ParentDescriptor As SECURITY_DESCRIPTOR, ByRef CreatorDescriptor As SECURITY_DESCRIPTOR, ByRef NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As Long, ByRef GenericMapping As GENERIC_MAPPING) As Long
'Private Declare Function apiGetSecurityInfo Lib "ADVAPI32" Alias "GetSecurityInfo" (ByVal Handle As Long, ByRef ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ByRef ppsidOwner As Sid, ByRef ppsidGroup As Sid, ByRef ppDacl As pAcl, ByRef ppSacl As pAcl, ByRef ppSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
'Private Declare Function apiGetUserObjectSecurity Lib "user32" Alias "GetUserObjectSecurity" (ByVal hObj As Long, ByRef pSIRequested As Long, ByRef pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
'Private Declare Function apiSetUserObjectSecurity Lib "user32" Alias "SetUserObjectSecurity" (ByVal hObj As Long, ByRef pSIRequested As Long, ByRef pSd As SECURITY_DESCRIPTOR) As Long
