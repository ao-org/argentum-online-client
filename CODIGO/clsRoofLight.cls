VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoofLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase para almacenar la luz del techo: alpha y color
' Es necesario porque VB solo acepta tipos primitivos u objetos de clase
' en colecciones y diccionarios, no acepta estructuras
Public Alpha As Single
Public Color As Long
