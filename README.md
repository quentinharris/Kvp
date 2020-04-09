# Kvp

A C# library that implements a flexible typeless Dictionary (Key Value Pairs) for use in VBA.
The objective is remove some of the pain (particularly the ppopulating step) of using Collections and Scripting.Dictionaries in VBA.

I'm not a professional programmer so don't expect to see an epitomy of idiomatic c# code.  
Much of what I've achieved is by virtue of Cargo Cult programming and reading the Blog articles at RubberDuck.  
The code is provided as is with no pretentions to warranty or fitness for purpose.  Use at your own risk.  

I'm keen to learn and do better, so suggestions, comments and contributions are most welcome.

## Using a Kvp

1. Place the KvpClass library elements in a directory.  
2. From the VBA IDE use Tools.References and then browse to the directory.  
3. Select the KvpClass.tlb file
4. Check the new entry in the References list (Kvp: A flexible Key/Value pair dictionary object for VBA)
5. Click OK.  

IN VBA use 
```
Dim myKvp as Kvp
Set myKvp = New Kvp
```

### Testing

The installed library can be tested using RubberDuck unit testing facilities with the file VBA.TestKvp.bas

### Strongly typed Kvp

By default a Kvp will accept any legal object/primitive as a key or value.  

A strongly typed Kvp can be achieved using VBA.KvpWrapper.bas


