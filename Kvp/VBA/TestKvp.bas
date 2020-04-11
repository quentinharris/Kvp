Attribute VB_Name = "TestKvp"
Option Explicit
'@IgnoreModule
'@TestModule
'@Folder("VBASupport")


Private Assert                                  As Rubberduck.AssertClass
'Private Fakes                                   As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    'Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

''@TestInitialize
'Private Sub TestInitialize()
''this method runs before every test in the module.
'End Sub
'
'
''@TestCleanup
'Private Sub TestCleanup()
''this method runs after every test in the module.
'End Sub


'@TestMethod("Kvp")
Private Sub IsObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = New Kvp
    
    'Assert:
    Assert.AreEqual "Kvp", TypeName(myKvp)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub IsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    
    'Act:
    Set myKvp = New Kvp
    
    'Assert:
    Assert.AreEqual True, myKvp.IsEmpty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub IsNotEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = New Kvp
    myKvp.AddbyIndex "Hello World"
    'Assert:
    Assert.AreEqual True, myKvp.IsNotEmpty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Count()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = New Kvp
    myKvp.AddbyIndex "Hello World"
    'Assert:
    Assert.AreEqual 1&, myKvp.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndex_NoStarterIndex()
    On Error GoTo TestFail

    'Arrange:
    Dim myArray As Variant
    myArray = Split("1,2,3,4,5,6,7,8,9", ",")
    
    'Act:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myItem As Variant
    For Each myItem In myArray
        myKvp.AddbyIndex myItem
    Next
    
    'Assert:
    Dim myResult As Boolean: myResult = True
    Dim myPair As Variant
    For Each myPair In myKvp
        myResult = myResult And (CStr(myPair.Key + 1) = myPair.Value)
    Next
    
    
    Assert.IsTrue myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndex_StarterIndexAsLong_5()
    On Error GoTo TestFail

    'Arrange:
    Dim myArray As Variant
    myArray = Split("1,2,3,4,5,6,7,8,9", ",")
    
    'Act:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.SetFirstIndexAsLong 5&
    Dim myItem As Variant
    For Each myItem In myArray
        myKvp.AddbyIndex myItem
    Next
    
    'Assert:
    Dim myResult As Boolean: myResult = True
    Dim myPair As Variant
    For Each myPair In myKvp
        myResult = myResult And (CStr(myPair.Key) = myPair.Value)
    Next
    
    
    Assert.AreEqual "1", myKvp.GetItem(5&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndex_StarterIndexAsString_Helloa()
    On Error GoTo TestFail

    'Arrange:
    Dim myArray As Variant
    myArray = Split("1,2,3,4,5,6,7,8,9", ",")
    
    'Act:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.SetFirstIndexAsString "Helloa"
    Dim myItem As Variant
    For Each myItem In myArray
        myKvp.AddbyIndex myItem
    Next
    
    'Assert:
    Dim myResult As Boolean: myResult = True
    Dim myPair As Variant
    For Each myPair In myKvp
        myResult = myResult And (CStr(myPair.Key) = myPair.Value)
    Next
    
    
    Assert.AreEqual "2", myKvp.GetItem("Hellob")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByIndexAsChars_DefaultStartIndex()
    On Error GoTo TestFail

    'Arrange:
    
    Dim myString As String
    myString = "Hello"
    
    Dim myExpectedArray(4) As String
    myExpectedArray(0) = "H"
    myExpectedArray(1) = "e"
    myExpectedArray(2) = "l"
    myExpectedArray(3) = "l"
    myExpectedArray(4) = "o"
    
    
    Dim myResultArray                            As Variant
    'Act:
    
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByIndexAsCharacters myString
    myResultArray = myKvp.GetValues
    'Assert:

    Assert.SequenceEquals myExpectedArray, myResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndexAsChars_StartIndexLong_5()
    On Error GoTo TestFail

    'Arrange:
    
    Dim myString As String
    myString = "Hello"
    
    Dim myExpectedArray As Variant
    myExpectedArray = Array(5&, 6&, 7&, 8&, 9&)
    
    Dim myResultArray  As Variant
    'Act:
    
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.SetFirstIndexAsLong 5&
    myKvp.AddByIndexAsCharacters myString
    myResultArray = myKvp.GetKeys
    'Assert:

    Assert.SequenceEquals myExpectedArray, myResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByIndexAsChars_StartIndexString_Helloa()
    On Error GoTo TestFail

    'Arrange:
    
    Dim myString As String
    myString = "Hello"
    
    Dim myExpectedArray As Variant
    myExpectedArray = Split("Helloa,Hellob,Helloc,Hellod,Helloe", ",")
    
    Dim myResultArray  As Variant
    'Act:
    
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.SetFirstIndexAsString "Helloa"
    myKvp.AddByIndexAsCharacters myString
    myResultArray = myKvp.GetKeys
    'Assert:

    Assert.SequenceEquals myExpectedArray, myResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndexFromCollection_DefaultStartIndex()
' An equivalent method is not required for Scripting.Dictionaries because
' we have AddByKeyFromArrays
    On Error GoTo TestFail

    'Arrange:
    
    Dim myArray As Variant: myArray = Array("Hello", True, 42, 3.142)
    
    Dim myCollection As Collection: Set myCollection = New Collection
    myCollection.Add myArray(0)
    myCollection.Add myArray(1)
    myCollection.Add myArray(2)
    myCollection.Add myArray(3)
    
    Dim myResult_Array                              As Variant
    Dim myKvp As Kvp: Set myKvp = New Kvp
    'Act:
    
    myKvp.AddByIndexFromArray myCollection
    myResult_Array = myKvp.GetValues
    'Assert:

    Assert.SequenceEquals myArray, myResult_Array

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByIndexFromArray_DefaultStartIndex()
    On Error GoTo TestFail

    'Arrange:
    Dim myKvp                                       As Kvp
    Dim myArray                                     As Variant
    Dim myResult_Array                              As Variant
    'Act:
    myArray = Array("Hello", True, 42, 3.142)
    
    Set myKvp = New Kvp
    myKvp.AddByIndexFromArray myArray
    myResult_Array = myKvp.GetValues
    'Assert:

    Assert.SequenceEquals myArray, myResult_Array

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndexFromArray_StartIndexIsLong_5()
    On Error GoTo TestFail

    'Arrange:
    Dim myArray As Variant
    myArray = Array("Hello", True, 42, 3.142)
    
    Dim myExpectedKeys As Variant
    myExpectedKeys = Array(5&, 6&, 7&, 8&)
    
    Dim myResultKeys As Variant
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.SetFirstIndexAsLong 5&
    myKvp.AddByIndexFromArray myArray
    
    myResultKeys = myKvp.GetKeys
    'Assert:

    Assert.SequenceEquals myExpectedKeys, myResultKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByIndexFromArray_StartIndexIsString_Helloa()
    On Error GoTo TestFail

    'Arrange:
    Dim myArray As Variant
    myArray = Array("Hello", True, 42, 3.142)
    
    Dim myExpectedKeys As Variant
    myExpectedKeys = Split("Helloa,Hellob,Helloc,Hellod", ",")
    
    Dim myResultKeys As Variant
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.SetFirstIndexAsString "Helloa"
    myKvp.AddByIndexFromArray myArray
    
    myResultKeys = myKvp.GetKeys
    'Assert:

    Assert.SequenceEquals myExpectedKeys, myResultKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod("Kvp")
Private Sub Add_ByKey_Long_101()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 22"
    myKvp.AddByKey Key:=25&, Value:="Hello World 25"
    myKvp.AddByKey Key:=31&, Value:="Hello World 31"
    myKvp.AddByKey Key:=101&, Value:="Hello World 101"
    myKvp.AddByKey Key:=2500&, Value:="Hello World 2500"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual "Hello World 101", myKvp.GetItem(101&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKey_String_Helloc()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:="Helloa", Value:="Hello World 22"
    myKvp.AddByKey Key:="Hellob", Value:="Hello World 25"
    myKvp.AddByKey Key:="Helloc", Value:="Hello World 31"
    myKvp.AddByKey Key:="Hellod", Value:="Hello World 101"
    myKvp.AddByKey Key:="Helloe", Value:="Hello World 2500"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual "Hello World 31", myKvp.GetItem("Helloc")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKeyFromArrays_KeysMatch()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKeyArray As Variant
    myKeyArray = Split("Key1,Key2,Key3,Key4,Key5", ",")
    
    Dim myValueArray As Variant
    myValueArray = Split("Val1,Val2,Val3,Val4,Val5", ",")
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddbyKeyFromArrays myKeyArray, myValueArray
   
    'Assert:
    
    Assert.SequenceEquals myKeyArray, myKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKeyFromArrays_ValuesMatch()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKeyArray As Variant
    myKeyArray = Split("Key1,Key2,Key3,Key4,Key5", ",")
    
    Dim myValueArray As Variant
    myValueArray = Split("Val1,Val2,Val3,Val4,Val5", ",")
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddbyKeyFromArrays myKeyArray, myValueArray
   
    'Assert:
    
    Assert.SequenceEquals myValueArray, myKvp.GetValues

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByKeyFromTable_CopyCol1AsKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myArray(4, 4) As Long
    Dim myRow As Long
    For myRow = 0 To 4
        
        Dim myCOl As Long
        For myCOl = 0 To 4
        
            myArray(myRow, myCOl) = (myCOl + 2) * (myRow + 1)
            Debug.Print myArray(myRow, myCOl);
        Next
        Debug.Print
    Next
    Debug.Print
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myExpectedKeys As Variant
    myExpectedKeys = Array(2&, 4&, 6&, 8&, 10&)
    'Act:
    myKvp.AddByKeyFromTable myArray, CopyKeys:=True
    
    Dim myPair As Variant
    For Each myPair In myKvp
        Debug.Print myPair.Value.GetValuesAsString
    Next
   
    'Assert:
    
    Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKeyFromTable_byColumnCopyCol1AsKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myArray(4, 4) As Long
    Dim myRow As Long
    For myRow = 0 To 4
        
        Dim myCOl As Long
        For myCOl = 0 To 4
        
            myArray(myRow, myCOl) = (myCOl + 2) * (myRow + 1)
            'Debug.Print myArray(myRow, myCOl);
            
        Next
        'Debug.Print
    Next
    'Debug.Print
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myExpectedKeys As Variant
    myExpectedKeys = Array(2&, 3&, 4&, 5&, 6&)
    
    
    'Act:
    myKvp.AddByKeyFromTable myArray, CopyKeys:=True, byColumn:=True
'   Dim myPair As Variant
'    For Each myPair In myKvp
'        Debug.Print myPair.Value.GetValuesAsString
'    Next
    'Assert:
    
    Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKeyFromTable_NoCopyCol1Keys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myArray(4, 4) As Long
    Dim myRow As Long
    For myRow = 0 To 4
        
        Dim myCOl As Long
        For myCOl = 0 To 4
        
            myArray(myRow, myCOl) = (myCOl + 2) * (myRow + 1)
            
        Next
        
    Next
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myExpectedKeys As Variant
    myExpectedKeys = Array(2&, 4&, 6&, 8&, 10&)
    
    'Act:
    myKvp.AddByKeyFromTable myArray
   
    'Assert:
    Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Add_ByKeyFromTable_byColumn_NoCopyCol1Keys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myArray(4, 4) As Long
    Dim myRow As Long
    For myRow = 0 To 4
        
        Dim myCOl As Long
        For myCOl = 0 To 4
        
            myArray(myRow, myCOl) = (myCOl + 2) * (myRow + 1)
            
        Next
        
    Next
    
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myExpectedKeys As Variant
    myExpectedKeys = Array(2&, 3&, 4&, 5&, 6&)
    
    'Act:
    myKvp.AddByKeyFromTable myArray, byColumn:=True
   
    'Assert:
    Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("Kvp")
Private Sub HoldsKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    
    'Assert:
    Assert.AreEqual True, myKvp.HoldsKey(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub HoldsValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    
    'Assert:
    Assert.AreEqual True, myKvp.HoldsValue("Hello World")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub LacksValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    
    'Assert:
    Assert.AreEqual True, myKvp.LacksValue("Hello There")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub LacksKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    'Act:
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    
    'Assert:
    Assert.AreEqual True, myKvp.LacksKey(80&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




'@TestMethod("Kvp")
Private Sub GetKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myKvp_keys(2)                       As Variant
    Dim myResult_Key                        As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myKvp_keys(0) = 22&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 25&
    
    myResult_Key = myKvp.GetKey("Hello World 2")
    'Assert:
    Assert.AreEqual myKvp_keys(1), myResult_Key
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myKvp_keys(2)                       As Long
    Dim myResult_Keys                       As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myKvp_keys(0) = 22&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 25&
    
    myResult_Keys = myKvp.GetKeys
    'Assert:
    Assert.SequenceEquals myKvp_keys, myResult_Keys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeysAsString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myKvp_keys(2)                       As Long
    Dim myResult                            As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myKvp_keys(0) = 22&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 25&
    Dim myExpected As Variant
    myExpected = "22,23,25"
    myResult = myKvp.GetKeysAsString
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeysAsStringAscending()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myResult                            As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    Dim myExpected As Variant
    myExpected = "22,23,25"
    myResult = myKvp.GetKeysAsString
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeysAsStringDecending()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myResult                            As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    
    Dim myExpected As Variant
    myExpected = "25,23,22"
    myResult = myKvp.GetKeysAsStringDescending
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeysAscending()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myKvp_keys(2)                       As Long
    Dim myResult_Keys                       As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=25&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=22&, Value:="Hello World 3"
    
    myKvp_keys(0) = 22&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 25&
    
    myResult_Keys = myKvp.GetKeysAscending
    'Assert:
    Assert.SequenceEquals myKvp_keys, myResult_Keys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeysDescending()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As Kvp
    Dim myKvp_keys(2)                       As Long
    Dim myResult_Keys                       As Variant

    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myKvp_keys(0) = 25&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 22&
    
    myResult_Keys = myKvp.GetKeysDescending
    'Assert:
    Assert.SequenceEquals myKvp_keys, myResult_Keys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetFirst()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByKey Key:=23&, Value:="Hello World 1"
    myKvp.AddByKey Key:=25&, Value:="Hello World 2"
    myKvp.AddByKey Key:=22&, Value:="Hello World 3"
    
    Dim myResult As Variant
    'Act:
    
    ' NB If we don't use the Set statement then the Value
    ' only is assigned to myResult because Value is the default property
    Set myResult = myKvp.GetFirst
    
    'Assert:
    Assert.AreEqual CVar(23&), myResult.Key
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub GetLast()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByKey Key:=23&, Value:="Hello World 1"
    myKvp.AddByKey Key:=25&, Value:="Hello World 2"
    myKvp.AddByKey Key:=22&, Value:="Hello World 3"
    Dim myResult As Variant
    
    'Act:
    Set myResult = myKvp.GetLast
    
    'Assert:
    Assert.AreEqual CVar(22&), myResult.Key
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    ' Dynamicops integers rather than long
    Dim myitems(2)                          As String
    Dim myKvp_items()                       As Variant
    'Act:
    Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myitems(0) = "Hello World 1"
    myitems(1) = "Hello World 2"
    myitems(2) = "Hello World 3"
    myKvp_items = myKvp.GetValues
    'Assert:
    Assert.SequenceEquals myitems, myKvp_items

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetValuesAsString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    Dim myExpected As Variant
    myExpected = "Hello World 1,Hello World 2,Hello World 3"
    'Act:
    Dim myKvpItems  As Variant
    myKvpItems = myKvp.GetValuesAsString
    
    'Assert:
    Assert.AreEqual myExpected, myKvpItems

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub KeyAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByKey Key:=23&, Value:="Hello World 1"
    myKvp.AddByKey Key:=25&, Value:="Hello World 2"
    myKvp.AddByKey Key:=22&, Value:="Hello World 3"
    Dim myResult As Variant
    
    'Act:
    myResult = myKvp.KeyAt(1)
    
    'Assert:
    Assert.AreEqual 25&, myResult
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub Cohorts_AllAandBOnly()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1 As Kvp: Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Dim myKvp2 As Kvp: Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    ' Cohort All Unique Keys: All keys in myKvp1 and keys in myKvp2 which are not in myKvp1
    Dim myResult_Keys  As Variant
    myResult_Keys = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&)
    
    'Act:
    Dim myResult As Kvp
    Set myResult = myKvp1.Cohorts(myKvp2)
    
    Dim myCohortKeys As Variant
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_AllAandBOnly).GetKeys
   
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys

    Set myKvp1 = Nothing
    Set myKvp2 = Nothing
    myResult.SetItem 0&, Nothing
    myResult.SetItem 1&, Nothing
    myResult.SetItem 2&, Nothing
    myResult.SetItem 3&, Nothing
    myResult.SetItem 4&, Nothing
    myResult.SetItem 5&, Nothing
    Set myResult = Nothing
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_AandBDifferentValues()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1 As Kvp: Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Dim myKvp2 As Kvp: Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    ' Cohort All Unique Keys: All keys in myKvp1 and keys in myKvp2 which are not in myKvp1
    Dim myResult_Keys  As Variant
    myResult_Keys = Array(3&)
    
    
    'Act:
    Dim myResult As Kvp
    Set myResult = myKvp1.Cohorts(myKvp2)
    
    Dim myCohortKeys As Variant
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_AandBDifferentValues).GetKeys
   
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys

    Set myKvp1 = Nothing
    Set myKvp2 = Nothing
    myResult.SetItem 0&, Nothing
    myResult.SetItem 1&, Nothing
    myResult.SetItem 2&, Nothing
    myResult.SetItem 3&, Nothing
    myResult.SetItem 4&, Nothing
    myResult.SetItem 5&, Nothing
    Set myResult = Nothing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_InAorInB()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    Dim myKvp2                                 As Kvp
    Dim myResult_Keys(3)                       As Long
    Dim myResult                             As Kvp
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_Keys(0) = 4&
    myResult_Keys(1) = 5&
    myResult_Keys(2) = 7&
    myResult_Keys(3) = 8&
    
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_InAorInB).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts__AandBSameValue()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    Dim myKvp2                                 As Kvp
    Dim myResult_Keys(3)                       As Long
    Dim myResult                             As Kvp
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_Keys(0) = 1&
    myResult_Keys(1) = 2&
    myResult_Keys(2) = 3&
    myResult_Keys(3) = 6&
    
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_AandBSameValues).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts__Aonly()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    Dim myKvp2                                 As Kvp
    Dim myResult_Keys(1)                  As Long
    Dim myResult                             As Kvp
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_Keys(0) = 4&
    myResult_Keys(1) = 5&
  
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_Aonly).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_Bonly()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    Dim myKvp2                                 As Kvp
    Dim myResult_Keys(1)                       As Long
    Dim myResult                             As Kvp
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New Kvp
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_Keys(0) = 7&
    myResult_Keys(1) = 8&
  
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.GetItem(KvpClass.Cohort_Bonly).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_Keys, myCohortKeys


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")

Private Sub Mirror()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                     As Kvp
    Dim myKvp2                                     As Kvp
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Set myKvp2 = myKvp1.Mirror
    
    'Assert:
    Assert.SequenceEquals myKvp1.GetKeys, myKvp2.Item(1&).GetValues

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub ItemsAreUnique()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    
    
    'Assert:
    Assert.AreEqual True, myKvp1.ValuesAreUnique

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub PullFirst()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As KVPair
    Set myResult = myKvp1.PullFirst
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(22&) And myResult.Value = "Hello World 1"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub PullLast()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As KVPair
    Set myResult = myKvp1.PullLast
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(27&) And myResult.Value = "Hello World 5"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub PullAny()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As Variant
    Set myResult = myKvp1.Pull(25&)
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(25&) And myResult.Value = "Hello World 3"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub RemoveFirst()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As String
    myKvp1.RemoveFirst
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub RemoveLast()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As String
    myKvp1.RemoveLast
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(27&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub RemoveAny()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As Kvp
    
    'Act:
    Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    Debug.Print myKvp1.GetKeysAsString
    Dim myResult As String
    myKvp1.Remove 25&
    Debug.Print myKvp1.GetKeysAsString
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(25&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub RemoveAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp.AddByKey Key:=27&, Value:="Hello World 5"
    
    'Act:
    myKvp.RemoveAll
    
    'Assert:
    Assert.AreEqual 0&, myKvp.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub SubSetByKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1 As Kvp: Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim mySubSetKeys As Variant
    mySubSetKeys = Array(22&, 25&, 26&)

    Dim mySubSetKvp As Kvp
    
    'Act:
    Set mySubSetKvp = myKvp1.SubSetByKeys(mySubSetKeys)
    
    'Assert:
    Assert.SequenceEquals mySubSetKeys, mySubSetKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub SubSetByKey_IgnoreKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1 As Kvp: Set myKvp1 = New Kvp
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim mySubSetKeys As Variant
    mySubSetKeys = Array(22&, 24&, 25&, 26&, 80&)

    Dim myResultKeys As Variant
    myResultKeys = Array(22&, 25&, 26&)

    Dim mySubSetKvp As Kvp
    
    'Act:
    Set mySubSetKvp = myKvp1.SubSetByKeys(mySubSetKeys)
    
    'Assert:
    Assert.SequenceEquals myResultKeys, mySubSetKvp.GetKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
