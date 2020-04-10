using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VBALibrary
{
    /// <summary>
    /// Enum for accessing the kvp structure returned by Method Cohorts
    /// </summary>

    public enum Cohort
    {
        /// <summary>1 = the keys in A plus keys in B that are not in A</summary>
        AllAandBOnly = 1,

        /// <summary>2 = the Keys from B in A where B has a different value to A</summary>
        AandBDifferentValues = 2,

        /// <summary>3 = the keys that are only in A and only in B</summary>
        InAorInB = 3,

        /// <summary>4 = the keys that are inA and  B </summary>
        AandBSameValues = 4,

        /// <summary>5 = the keys in A only   </summary>
        Aonly = 5,

        /// <summary>6 = the keys in B only</summary>
        Bonly = 6
    }

    /// <summary>
    /// Kvp is a C# class for VBA which implements a Key/Value Dictionary
    /// The object is a morer flexible version of the Scripting.Dictionary
    /// </summary>
    //[Guid("6DC1808F-81BA-4DE0-9F7C-42EA11621B7E")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IKvp
    {
        ///// <summary>
        ///// Returns/Sets the "Value" specified by "Key" (i) of a Key/Value Pair
        ///// </summary>
        ///// <param name="Key"></param>
        ///// <returns>Type used in Set statement (C# dynamic)</returns>
        //dynamic this[dynamic Key] { get; set; }

        dynamic GetItem(dynamic Key);

        void SetItem(dynamic Key, dynamic Value);

        IEnumerator GetEnumerator();

        void SetFirstIndexAsLong(int Index = 0);

        void SetFirstIndexAsString(string Index = "a");

        /// <summary>
        /// Adds "Value" to the Kvp using an integer (VBA Long) Key.
        /// The integer key is is started at 0
        /// </summary>
        /// <param name="Value"></param>
        void AddByIndex(dynamic Value);

        // void AddByIndexFromEnumerable(dynamic ipEnumerable);

        void AddByIndexAsBytes(string ipString);

        /// <summary>
        /// Populates this Kvp using AddByIndex for each character in the string
        /// </summary>
        /// <param name="ipString"></param>
        void AddByIndexAsLetters(string ipString);

        /// <summary>
        /// Pupulates this Kvp using AddByIndex for each substring in ipString delineated by ipSeparator
        /// </summary>
        /// <param name="ipString"></param>
        /// <param name="ipSeparator"></param>
        //void AddByIndexAsSubStr(string ipString, string ipSeparator = ",");

        /// <summary>
        /// Pupulates a Kvp using AddByIndex for each array item
        /// </summary>
        /// <param name="this_array"></param>
        void AddByIndexFromArray(dynamic ipArray);

        void AddByIndexFromCollection(dynamic ipCollection);

        /// <summary>
        /// Adds "Value" to the Kvp with a key pf "Key"
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="Value"></param>
        void AddByKey(dynamic Key, dynamic Value);

        /// <summary>
        /// Adds array to the Kvp with ujsing array[x,1] as the key and array[x,0..end] as the value
        /// </summary>
        //void AddByKeyFromArray(dynamic ipArray);

        void AddByKeyFromArrays(dynamic KeyArray, dynamic ValueArray);

        void AddByKeyFromTable(dynamic table, bool CopyKeys = true, bool byColumn = false);

        /// <summary>
        /// Groups the keys of the two Kvp
        /// </summary>
        /// <param name="ArgKvp"></param>
        /// <returns>An array of 6 Kvp
        /// keys in a {1,2,3,4,5,6}
        /// keys in b {1,2,3,6,7,8}
        /// 1 = the keys in A plus keys in B that are not shared            {1,2,3( from A),4,5,6,7,8}
        /// 2 = the Keys from B in A where B has a different value to A     {3( from B) if value is different}
        /// 3 = the keys that are only in A and only in B                   {4,5,7,8}
        /// 4 = the keys that are in A and  B                               {1,2,3,6}
        /// 5 = the keys in A only                                          {4,5}
        /// 6 = the keys in B only                                          {7,8}
        /// </returns>
        Kvp Cohorts(dynamic ArgKvp);

        /// <summary>
        /// The number of key/vaue pairs in the Kvp
        /// </summary>
        /// <returns>Long</returns>
        int Count();

        /// <summary>
        /// Returns a shallow copy of the Kvp
        /// </summary>
        /// <returns>New kvp as a copy of the old kvp</returns>
        Kvp Clone();

        // IEnumerator GetEnumerator();

        /// <summary>
        /// Gets the "Key" for the first ocurrence of "Value" in the Kvp.
        /// </summary>
        /// <param name="Value"></param>
        /// <returns>Key</returns>
        dynamic GetKey(dynamic Value);

        /// <summary>
        /// Returns a variant array of the Keys of the Kvp
        /// </summary>
        /// /// <returns>Variant Array</returns>
        dynamic GetKeys { get; }

        dynamic GetKeysAscending();

        /// <summary>
        /// Concatenates the keys to a string
        /// Will fall over is an element is not a primitive
        /// </summary>summary>
        /// <returns>Keys seperated by seperator</returns>
        string GetKeysAsString(string Separator = ","); //string ipSeparator = ",");

        string GetKeysAsStringAscending(string Separator = ",");

        dynamic GetKeysDescending();

        string GetKeysAsStringDescending(string Separator = ",");

        /// <summary>
        /// Gets the 0th element of the Kvp
        /// </summary>
        /// <returns>Value at lbound postion</returns>
        dynamic GetFirst();

        /// <summary>
        /// Gets the element at the upper bound of MyKvp
        /// This may not be the maximum integer key
        /// </summary>
        /// <returns>Value of the element at the upper bound</returns>
        dynamic GetLast();

        /// <summary>
        /// Returns a variant array of the values of the Kvp
        /// </summary>
        /// <returns>Variant Array</returns>
        dynamic GetValues { get; }

        /// <summary>
        /// Concatenates the values to a string
        /// Will fall over is an element is not a primitive
        /// </summary>summary>
        /// <returns>values seperated by seperator</returns>
        string GetValuesAsString(string Separator = ","); //string ipSeparator = ",");

        /// <summary>
        /// True if the "Key" exists in the keys of the Kvp
        /// </summary>
        /// <param name="Key"></param>
        /// <returns>Boolean</returns>
        bool HoldsKey(dynamic Key);

        /// <summary>
        /// True if the "Value" exists in the values of the Kvp
        /// </summary>
        /// <param name="Value"></param>
        /// <returns>Boolean</returns>
        bool HoldsValue(dynamic Value);

        /// <summary>
        /// True if the Kvp holds 0 key/value pairs
        /// </summary>
        /// <returns>Boolean</returns>
        bool IsEmpty();

        /// <summary>
        /// True if the Kvp holds one or more key/value pairs
        /// </summary>
        /// <returns>Boolean</returns>
        bool IsNotEmpty();

        /// <summary>
        /// Return the key at corresponding Index of the Index
        /// of the Value in the values array
        /// </summary>
        /// <param name="Index"></param>
        /// <returns>key at Index of Value in Values array</returns>
        dynamic KeyAt(int Index);

        IEnumerator KvpEnum();

        /// <summary>
        /// True is the "Key" is not found in the keys of the Kvp
        /// </summary>
        /// <param name="Key"></param>
        /// <returns>Boolean</returns>
        bool LacksKey(dynamic Key);

        /// <summary>
        /// True if the "Value" is not found in the values of the Kvp
        /// </summary>
        /// <param name="Value"></param>
        /// <returns>Boolean</returns>
        bool LacksValue(dynamic Value);

        /// <summary>
        /// Reverses the Key/Value pairs in a Kvp
        /// </summary>
        /// <returns>New Kvp where:
        ///     Kvp.Value(1) = Kvp of Unique values as Value/Key pairs
        ///     Kvp.Value(2) = Kvp of Non unique values as original Key/Value pairs</returns>
        Kvp Mirror();

        /// <summary>
        /// Gets the value of the key then deletes the element from MyKvp
        /// </summary>
        /// <param name="ipKey"></param>
        /// <returns>Value at key</returns>
        dynamic Pull(dynamic ipKey);

        /// <summary>
        /// get the value of the key at the lower bound of array keys
        /// deletes the elemnt
        /// </summary>
        /// <returns>value at key of lower bound element</returns>
        dynamic PullFirst();

        /// <summary>
        /// get the value of the key at the upper bound of array keys
        /// deletes the element
        /// </summary>
        /// <returns>value at key of upper bound element</returns>
        dynamic PullLast();

        /// <summary>
        /// Removes the Key/Value pair spacified by "Key" from the Kvp
        /// </summary>
        //void Remove(dynamic Key);

        /// <summary>
        /// deletes the element at th key of the lower bound of array MyKvp.keys
        /// </summary>
        void RemoveFirst();

        /// <summary>
        /// deletes the element with the key at the upper bound of array MyKvp.keys
        /// </summary>
        void RemoveLast();

        /// <summary>
        /// Removes all Key/Value pairs from the Kvp
        /// </summary>
        void RemoveAll();

        /// <summary>
        /// Returns true if the Values in Kvp are unique.
        /// </summary>
        /// <returns>Boolean</returns>
        bool ValuesAreUnique();

        /// <summary>
        /// Returns true if the Values in Kvp are not unique.
        /// </summary>
        /// <returns>Boolean</returns>
        bool ValuesAreNotUnique();
    }
}