Attribute VB_Name = "mType"
'Downloaded and modified from http://stackoverflow.com/questions/1847411/control-properties-in-visual-basic-6
Public Enum EPType
    ReadableProperties = 2
    WriteableProperties = 4
End Enum

Public Function EnumerateProperties(pObject As Object, pType As EPType) As Variant
    Dim rArray() As String
    Dim iVal As Long
    Dim TypeLib As TLI.InterfaceInfo
    Dim Prop As TLI.MemberInfo
    On Error Resume Next
    ReDim rArray(0) As String
    Set TypeLib = TLI.InterfaceInfoFromObject(pObject)
    For Each Prop In TypeLib.Members
        If Prop.InvokeKind = pType Then
            iVal = UBound(rArray)
            rArray(iVal) = Prop.Name
            ReDim Preserve rArray(iVal + 1) As String
        End If
    Next
    ReDim Preserve rArray(UBound(rArray) - 1) As String
    EnumerateProperties = rArray
End Function

Public Function PropertyExist(pObject As Object, ByVal _
    PropertyName As String, pType As EPType) As Boolean
    Dim Item As Variant
    PropertyExist = False
    For Each Item In EnumerateProperties(pObject, pType)
        If Item = PropertyName Then
            PropertyExist = True
            Exit For
        End If
    Next
End Function

