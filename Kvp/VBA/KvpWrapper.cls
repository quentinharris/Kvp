VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'@IgnoreModule
'@Folder("Kvp")
'@Exposed
'
'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'Replace KvpWrapper with the Class name
'Replace KeyType with your actual type. For a non typed Kvp use Variant
'Replace ValueType with your actual Type. For a non typed Kvp use Variant
'
'
'Private Const DefaultStringSeparator As String = ","
'
'Private Type State
'
'   Host                                         As Kvp
'
'End Type
'
'Private s                                       As State
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Class Admin
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'Private Sub Class_Initialize()
'
'    Set s.Host = New Kvp
'
'End Sub
'
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Implementation - stuff with wrinkles
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'@DefaultMember
'Public Property Get Item(ByVal Key As KeyType) As ValueType
'    Item = s.Host.GetItem(Key)
'    Set Item = s.Host.GetItem(Key)
'End Property
'
' Select Let or Set depending on the Value Type
'Public Property Let Item(ByVal Key As KeyType, ByVal Value As ValueType)
'    s.Host.SetItem Key, Value
'End Property
'
'Public Property Set Item(ByVal Key As KeyType, ByVal Value As ValueType)
'    s.Host.SetItem Key, Value
'End Property
'
'@Enumerator
'Public Property Get NewEnum() As Variant
'     KvpEnum returns a KVPair but the vagaries of VBA means we have to assign to a Variant
'     Set NewEnum = s.HostKvp.KvpEnum
'
'End Property
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Implementation - just passthoughs
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'Public Sub AddbyIndex(ByVal Value As ValueType)
'    s.Host.AddbyIndex Value
'End Sub
'
'Public Sub AddByIndexAsByte(ByVal ValueString As String)
'    s.Host.AddByIndexAsBytes ValueString
'End Sub
'
'Public Sub AddByIndexAsCharacters(ByVal ValueString As String)
'    s.Host.AddByIndexAsCharacters ValueString
'End Sub
'
'Public Sub AddByIndexFromArray(ByVal ValueArray As Variant)
'    s.Host.AddByIndexFromArray ValueArray
'End Sub
'
'Public Sub AddByIndexFromCollection(ByVal ValueCollection As Collection)
'    s.Host.AddByIndexFromCollection ValueCollection
'End Sub
'
'Public Sub AddByKey(ByVal Key As KeyType, ByVal Value As ValueType)
'    s.Host.AddByKey Key, Value
'End Sub
'
'Public Sub AddbyKeyFromArrays(ByVal KeyArray As Variant, ByVal ValueArray As Variant)
'    s.Host.AddbyKeyFromArrays KeyArray, ValueArray
'End Sub
'
'Public Sub AddByKeyFromTable(ByVal TableArray As Variant, Optional CopyKeys = False, Optional byColumn = False)
'    s.Host.AddByKeyFromTable Table:=TableArray, CopyKeys:=CopyKeys, byColumn:=byColumn
'End Sub
'
'Public Function Clone() As KvpWrapper
'    Set Clone = s.Host.Clone
'End Function
'
'Public Function Cohorts(ByVal KvpB As Object) As Kvp
'    Set Cohorts = s.Host.Cohorts(KvpB)
'End Function
'
'Public Function Count() As Long
'    Count = s.Host.Count
'End Function
'
'Public Function GetFirst() As KVPair
'    Set GetFirst = s.Host.GetFirst
'End Function
'
'Public Function GetKey(ByVal Value As ValueType) As KeyType
'    GetKey = s.Host.GetKey(Value)
'End Function
'
'Public Function GetKeys() As Variant
'    GetKeys = s.Host.GetKeys
'End Function
'
'Public Function GetKeysAscending() As Variant
'    GetKeys = s.Host.GetKeysAscending
'End Function
'
'Public Function GetKeyAsString(Optional ByVal Separator As String = DefaultStringSeparator) As String
'    GetKeysAsString = s.Host.GetKeysAsString(Separator)
'End Function
'
'Public Function GetKeyAsStringAscending(Optional ByVal Separator As String = DefaultStringSeparator) As String
'    GetKeysAsStringAscending = s.Host.GetKeysAscending(Separator)
'End Function
'
'Public Function GetKeyAsStringDescending(Optional ByVal Separator As String = DefaultStringSeparator) As String
'    GetKeysAsString = s.Host.GetKeysAsStringDescending(Separator)
'End Function
'
'Public Function GetKeysDescending() As Variant
'    GetKeysDescending = s.Host.GetKeysDescending
'End Function
'
'Public Function GetLast() As KVPair
'    Set GetLast = s.Host.GetLast
'End Function
'
'Public Function GetValues() As Variant
'    GetValues = s.Host.GetValues
'End Function
'
'Public Function GetValuesAsString() As String
'    GetValuesAsString = s.Host.GetValuesAsString
'End Function
'
'Public Function HoldsKey(ByVal Key As KeyType) As Boolean
'    HoldsKey = s.Host.HoldsKey(Key)
'End Function
'
'Public Function HoldsValue(ByVal Value As ValueType) As Boolean
'    HoldsValue = s.Host.HoldsValue
'End Function
'
'Public Function IsEmpty() As Boolean
'    IsEmpty = s.Host.IsEmpty
'End Function
'
'Public Function IsNotEmpty() As Boolean
'    IsNotEmpty = s.Host.IsNotEmpty
'End Function
'
'Public Function KeyAt(ByVal Index As Long) As KeyType
'    KeyAt = s.Host.KeyAt(Index)
'End Function
'
'Public Function LacksKey(ByVal Key As KeyType) As Boolean
'    LackKey = s.Host.LacksKey(Key)
'End Function
'
'Public Function LacksValue(ByVal Value As ValueType) As Boolean
'    LacksValue = s.Host.LacksValue(Value)
'End Function
'
'Public Function Mirror() As Kvp
'    Set Mirror = s.Host.Mirror
'End Function
'
'Public Function Pull(ByVal Key As KeyType) As KVPair
'    Set Pull = s.Host.Pull(Key)
'End Function
'
'Public Function PullFirst() As KVPair
'    Set PullFirst = s.Host.PullFirst
'End Function
'
'Public Function PullLast() As KVPair
'    Set PullLast = s.Host.PullLast
'End Function
'
'Public Sub Remove(ByVal Key As KeyType)
'    s.Host.Remove Key
'End Sub
'
'Public Sub RemoveFirst()
'    s.Host.RemoveFirst
'End Sub
'
'Public Sub RemoveLast()
'    s.Host.RemoveLast
'End Sub
'
'Public Sub SetFirstIndexAsLong(ByVal FirstIndex As Long)
'    s.Host.SetFirstIndexAsLong FirstIndex
'End Sub
'
'Public Sub SetFirstIndexAsString(ByVal FirstIndex As String)
'    s.Host.SetFirstIndexAsString FirstIndex
'End Sub
'
'Public Function SubSetByKeys(ByVal KeyArray As Variant) As Kvp
'    Set SubSetByKeys = s.Host.SubSetByKeys(KeyArray)
'End Function
'
'Public Function ValuesAreUnique() As Boolean
'    ValuesAreUnique = s.Host.ValuesAreUnique
'End Function
'
'Public Function ValuesAreNotUnique() As Boolean
'    ValuesAreNotUnique = s.Host.ValuesAreNotUnique
'End Function
