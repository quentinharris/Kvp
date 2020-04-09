# Kvp Methods and Properties

```
Dim myKvp as Kvp  ' or KvpClass.Kvp
Set myKvp = New Kvp
```
### AddByIndex

Adds the provided Item using the next available Key. 

```
myKvp.AddByIndex 42&
myKvp.AddByIndex "Hellow World"
myKvp.AddByIndex New Kvp
```
The Index will start at a value of 0 (VBA Long 0) if the starting point is not set via 
• SetFirstIndexAsLong, or
• SetFirstIndexAsString

For a standard Kvp there is no restriction on the type of key and consequently an Item added by .AddByIndex may be followed by AddbyKey with the key being specified as a string.

###AddByIndexAsByte

Adds the provided string as a sequence of integer values represented by the Ascii value of each character in the string

```
myKvp.AddByIndex "Hello World"
