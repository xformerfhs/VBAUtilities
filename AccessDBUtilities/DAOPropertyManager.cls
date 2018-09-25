VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAOPropertyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const ERR_PROPERTY_NOT_FOUND As Long = 3270

Public Property Let setProperty(ByRef obj As Object, _
                                ByRef propertyName As String, _
                                ByRef propertyType As Long, _
                                ByRef propertyValue As Variant)
   Dim properyList As DAO.Properties
   Dim aProperty As DAO.Property

   Set properyList = obj.Properties

   On Error Resume Next
   properyList(propertyName).Value = propertyValue

   If Err.Number <> 0 Then
      If Err.Number = ERR_PROPERTY_NOT_FOUND Then
         On Error GoTo 0

         Set aProperty = obj.CreateProperty(propertyName)

         With aProperty
            .Type = propertyType
            .Value = propertyValue
         End With

         properyList.Append aProperty

         Set aProperty = Nothing
      Else
         Err.Raise Err.Number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
      End If
   End If

   Set properyList = Nothing
End Property

Public Property Get getProperty(ByRef obj As Object, _
                                ByRef propertyName As String) As Variant
   Dim properyList As DAO.Properties
   Dim aProperty As DAO.Property

   Set properyList = obj.Properties

   On Error Resume Next
   getProperty = properyList(propertyName).Value
   
   If Err.Number <> 0 Then _
      getProperty = Null

   On Error GoTo 0
End Property

Public Property Get getPropertyWithDefault(ByRef obj As Object, _
                                           ByRef propertyName As String, _
                                           ByRef defaultValue As Variant) As Variant
   getPropertyWithDefault = Nz(Me.getProperty(obj, propertyName), defaultValue)
End Property
