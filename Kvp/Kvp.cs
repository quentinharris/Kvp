using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBALibrary
{
    //[Guid("434C844C-9FA2-4EC6-AB75-45D3013D75BE")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    //[ComDefaultInterface(typeof(IKvp))]
    //[ComSourceInterfaces(typeof(IKvp))] obsolete
    public class Kvp : IKvp, IEnumerable
    {
        private Dictionary<dynamic, dynamic> Host = new Dictionary<dynamic, dynamic>();

        // Item[x]: Not used in VBA due to the Intellisense/IDespatch issue
        public dynamic this[dynamic key]
        {
            get
            {
                return Host[key];
            }
            set
            {
                Host[key] = (dynamic)value;
            }
        }

        // for use by VBA to replace = Kvp.Item(x)
        public dynamic GetItem(dynamic Key)
        {
            return Host[Key];
        }

        // for use by VBA to replace Kvp.Item(x)=
        public void SetItem(dynamic Key, dynamic Value)
        {
            Host[Key] = Value;
        }

        // Allows a standard Kvp to be enumerated in VBA
        public IEnumerator GetEnumerator()
        {
            foreach (KeyValuePair<dynamic, dynamic> myPair in Host)
            {
                yield return new KVPair(myPair.Key, myPair.Value);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        // Allows a wrapped Kvp to be enumerated in VBA
        public IEnumerator KvpEnum()
        {
            foreach (KeyValuePair<dynamic, dynamic> myPair in Host)
            {
                yield return new KVPair(myPair.Key, myPair.Value);
            }
        }

        private dynamic _FirstIndex;

        public void SetFirstIndexAsLong(int Index = 0)
        {
            if (_FirstIndex != null)
            {
                throw new ArgumentNullException("Kvp:FirstIndex: The first index is already set as long {0}", _FirstIndex);
            }
            _FirstIndex = Index;
        }

        public void SetFirstIndexAsString(string Index = "a")
        {
            if (_FirstIndex != null)
            {
                throw new ArgumentNullException("Kvp:FirstIndex: The first index is already set as string {0}", _FirstIndex);
            }
            _FirstIndex = Index;
        }

        // An equivalent Method for Scripting.Dictionaries is not required bewcause
        // we have the AddbyKeyFromArrays Method
        public void AddByIndexFromCollection(dynamic ipCollection)
        {
            if (ipCollection is null)
            {
                throw new ArgumentNullException("Kvp:AddByIndexAsChars: The provided Collection was null");
            }
            foreach (dynamic myItem in ipCollection)
            {
                Host.Add(GetNextKvpKey(), myItem);
            }
        }

        public void AddByIndex(dynamic Value)
        {
            Host.Add(GetNextKvpKey(), Value);
        }

        public void AddByIndexFromArray(dynamic ipArray)
        {
            //dynamic[] myArray = (dynamic)ipArray;
            foreach (dynamic myItem in ipArray)
            {
                Host.Add(GetNextKvpKey(), myItem);
            }
        }

        public void AddByIndexAsByte(string ipString)
        {
            if (ipString.Length == 0)
            {
                throw new ArgumentNullException("Kvp:AddByIndexAsChars: The provided string was empty (Length=0)");
            }
            foreach (char myChar in ipString)
            {
                Host.Add(GetNextKvpKey(), myChar);
            }
        }

        public void AddByIndexAsLetters(string ipString)
        {
            if (ipString.Length == 0)
            {
                throw new ArgumentNullException("Kvp:AddByIndexAsChars: The provided string was empty (Length=0)");
            }
            foreach (char myChar in ipString)
            {
                Host.Add(GetNextKvpKey(), myChar.ToString());
            }
        }

        public void AddByKey(dynamic Key, dynamic Value)
        {
            Host.Add(Key, Value);
        }

        //public void AddByKeyFromArray(dynamic ipArray)
        //{
        //    int rowLower = ipArray.GetLowerBound(0);
        //    int rowUpper = ipArray.GetUpperBound(0);
        //    int colLower = ipArray.GetLowerBound(1);
        //    int colUpper = ipArray.GetUpperBound(1);

        //    for (int i = rowLower; i < rowUpper; i++)
        //    {
        //        Kvp thisRow = new Kvp();
        //        for (int j = colLower; j < colUpper; j++)
        //        {
        //            thisRow.AddByIndex(ipArray[i, j]);
        //        }
        //        MyKvp.Add(ipArray[i, 1], thisRow);
        //    }
        //}

        public void AddByKeyFromArrays(dynamic KeyArray, dynamic ValueArray)
        {
            int KeysCount = ((dynamic[])KeyArray).GetUpperBound(0);
            int ValuesCount = ((dynamic[])ValueArray).GetUpperBound(0);
            if (KeysCount != ValuesCount)
            {
                throw new ArgumentOutOfRangeException("Kvp:AddByKeyFromArrays: Mismatched array sizes.");
            }
            if (ValueArray.Rank > 2)
            {
                throw new ArgumentOutOfRangeException("Kvp:AddByKeyFromArrays: The Value array must be a 1D array.\n Use AddByIndexFromTable for 2D arrays");
            }
            for (int i = 0; i <= KeysCount; i++)
            {
                Host.Add(KeyArray[i], ValueArray[i]);
            }
        }

        public void AddByKeyFromTable(dynamic table, bool CopyKeys = false, bool byColumn = false)
        {
            const int RowDimension = 0;
            const int ColDimension = 1;
            if (table.Rank != 2)
            {
                throw new ArgumentException("Kvp:AddByKeyFromTable: Two dimensional array expected");
            }

            int rowFirst = table.GetLowerBound(RowDimension);
            int rowLast = table.GetUpperBound(RowDimension);
            int colFirst = table.GetLowerBound(ColDimension);
            int colLast = table.GetUpperBound(ColDimension);
            if (byColumn)
            {
                for (int ThisColumn = colFirst; ThisColumn <= colLast; ThisColumn++)
                {
                    Kvp currentRow = new Kvp();
                    if (!CopyKeys)
                    { rowFirst += 1; }
                    for (int ThisRow = rowFirst; ThisRow <= rowLast; ThisRow++)
                    {
                        currentRow.AddByIndex(table[ThisRow, ThisColumn]);
                    }
                    Host.Add(table[0, ThisColumn], currentRow);
                }
            }
            else
            {
                for (int ThisRow = rowFirst; ThisRow <= rowLast; ThisRow++)
                {
                    Kvp currentCol = new Kvp();
                    if (!CopyKeys)
                    { colFirst += 1; }
                    for (int ThisColumn = colFirst; ThisColumn <= colLast; ThisColumn++)
                    {
                        currentCol.AddByIndex(table[ThisRow, ThisColumn]);
                    }
                    Host.Add(table[ThisRow, 0], currentCol);
                }
            }
        }

        // May need to change Kvp to dynamic
        public Kvp Cohorts(dynamic ArgKvp)
        {
            Kvp ResultKvp = new Kvp();
            // VBA reports object not set error if the result kvps are not newed
            for (int i = 1; i <= 6; i++)
            {
                ResultKvp.AddByKey(i, new Kvp());
            }
            // Process Kvp A
            foreach (KeyValuePair<dynamic, dynamic> myPair in Host)
            {
                // A plus unique in B
                ResultKvp[(int)Cohort.AllAandBOnly].AddByKey(myPair.Key, myPair.Value);

                if (ArgKvp.LacksKey(myPair.Key))
                {
                    // In A only or in B only
                    ResultKvp[(int)Cohort.InAorInB].AddByKey(myPair.Key, myPair.Value);
                    // In A only
                    ResultKvp[(int)Cohort.Aonly].AddByKey(myPair.Key, myPair.Value);
                }
                else
                {
                    // In A and In B
                    ResultKvp[(int)Cohort.AandBSameValues].AddByKey(myPair.Key, myPair.Value);
                }
            }

            //Process Kvp B
            foreach (KVPair myPair in ArgKvp)
            {
                // B in A with different value
                if (!Host.ContainsKey(myPair.Key))
                {
                    ResultKvp[(int)Cohort.AllAandBOnly].AddByKey(myPair.Key, myPair.Value);
                    ResultKvp[(int)Cohort.InAorInB].AddByKey(myPair.Key, myPair.Value);
                    ResultKvp[(int)Cohort.Bonly].AddByKey(myPair.Key, myPair.Value);
                }
                else
                {
                    var myHost = Host[myPair.Key];
                    var myTest = myPair.Value;
                    if ((dynamic)myHost != (dynamic)myTest)
                    {
                        ResultKvp.GetItem((int)Cohort.AandBDifferentValues).AddByKey(myPair.Key, new Kvp());
                        ResultKvp.GetItem((int)Cohort.AandBDifferentValues).GetItem(myPair.Key).AddByIndex(myHost);
                        ResultKvp.GetItem((int)Cohort.AandBDifferentValues).GetItem(myPair.Key).AddByIndex(myPair.Value);
                    }
                }
            }
            return ResultKvp;
        }

        public int Count()
        {
            return Host.Count;
        }

        public Kvp Clone()
        {
            Kvp CloneKvp = new Kvp();
            foreach (KeyValuePair<dynamic, dynamic> myPair in Host)
            {
                CloneKvp.AddByKey(myPair.Key, myPair.Value);
            }
            return CloneKvp;
        }

        public dynamic GetKey(dynamic Value)
        {
            return Mirror()[1][Value];
        }

        public dynamic GetKeys
        {
            get
            {
                return Host.Keys.ToArray();
            }
        }

        public dynamic GetKeysAscending()
        {
            dynamic myKeys = Host.Keys.ToArray();
            Array.Sort(myKeys);
            return myKeys;
        }

        public string GetKeysAsString(string Separator = ",")
        {
            return Host.Keys.ToSeparatedString(Separator);
        }

        public string GetKeysAsStringAscending(string Separator = ",")
        {
            return String.Join(Separator, GetKeysAscending());
        }

        public string GetKeysAsStringDescending(string Separator = ",")
        {
            return String.Join<dynamic>(Separator, GetKeysDescending());
        }

        public dynamic GetKeysDescending()
        {
            dynamic myKeys = Host.Keys.ToArray();
            Array.Sort(myKeys);
            Array.Reverse(myKeys);
            return myKeys;
        }

        public dynamic GetFirst()
        {
            dynamic myPair = Host.First();
            return new KVPair(myPair.Key, myPair.Value);
        }

        public dynamic GetLast()
        {
            dynamic myPair = Host.Last();
            return new KVPair(myPair.Key, myPair.Value);
        }

        public dynamic GetValues
        {
            get
            {
                return Host.Values.ToArray();
            }
        }

        public string GetValuesAsString(string Separator = ",") //string ipSeparator = ",")
        {
            return Host.Values.ToSeparatedString(Separator);
        }

        public bool HoldsKey(dynamic Key)
        {
            return Host.ContainsKey(Key);
        }

        public bool HoldsValue(dynamic Value)
        {
            return Host.ContainsValue(Value);
        }

        public bool IsEmpty()
        {
            return Host.Count == 0;
        }

        public bool IsNotEmpty()
        {
            return !IsEmpty();
        }

        public dynamic KeyAt(int Index)
        {
            if (Index < 0 || Index >= Host.Count)
            {
                throw new ArgumentOutOfRangeException("Kvp:KeyAt: Index outside Count range");
            }
            List<dynamic> myKeys = Host.Keys.ToList();
            return myKeys[Index];
        }

        public bool LacksKey(dynamic Key)
        {
            return !HoldsKey(Key);
        }

        public bool LacksValue(dynamic Value)
        {
            return !HoldsValue(Value);
        }

        public Kvp Mirror()
        {
            Kvp MyResult = new Kvp();
            MyResult.AddByKey((int)1, new Kvp());
            MyResult.AddByKey((int)2, new Kvp());
            foreach (KeyValuePair<dynamic, dynamic> my_pair in Host)
            {
                if (MyResult[1].LacksKey(my_pair.Value))
                {
                    MyResult[1].AddByKey(my_pair.Value, my_pair.Key);
                }
                else
                {
                    MyResult[2].AddByKey(my_pair.Key, my_pair.Value);
                }
            }
            return MyResult;
        }

        public dynamic Pull(dynamic ipKey)
        {
            dynamic myValue = Host[ipKey];
            Host.Remove(ipKey);
            return new KVPair(ipKey, myValue);
        }

        public dynamic PullFirst()
        {
            dynamic myPair = Host.First();
            Host.Remove(myPair.Key);
            return new KVPair(myPair.Key, myPair.Value);  //.Key;//, myPair.Value); ; ;
        }

        public dynamic PullLast()
        {
            dynamic myPair = Host.Last();
            Host.Remove(myPair.Key);
            return new KVPair(myPair.Key, myPair.Value);
        }

        public void Remove(dynamic Key)
        {
            Host.Remove(Key);
        }

        public void RemoveFirst()
        {
            this.PullFirst();
        }

        public void RemoveLast()
        {
            this.PullLast();
        }

        public void RemoveAll()
        {
            Host.Clear();
        }

        public bool ValuesAreUnique()
        {
            return Host.Count == Host.Values.Distinct().Count();
        }

        public bool ValuesAreNotUnique()
        {
            return !ValuesAreUnique();
        }

        private dynamic GetNextKvpKey()
        {
            //This is not an iterator function
            if (Host.Count == 0)
            {
                if (_FirstIndex == null)
                {
                    _FirstIndex = 0;
                    return _FirstIndex;
                }
                else
                {
                    return _FirstIndex;
                }
            }

            var myKey = Host.Keys.Last();
            if (IsNumericDataType(myKey))
            {
                return GetNextKvpKeyAsNumber(myKey);
            }

            if (myKey is string)
            {
                return GetNextKvpKeyAsString(myKey);
            }

            throw new InvalidOperationException("Kvp:Adding by Index: Key is not a key or number");
        }

        private dynamic GetNextKvpKeyAsNumber(dynamic ipKey)
        {
            while (Host.ContainsKey(ipKey))
            {
                ipKey++;
            }
            return ipKey;
        }

        private string GetNextKvpKeyAsString(string ipKey)
        {
            var myKey = ipKey;
            while (Host.ContainsKey(myKey))
            {
                myKey = IncrementStr(myKey);
            }
            return myKey;
        }

        private bool IsNumericDataType(dynamic ipValue)
        {
            switch (Type.GetTypeCode(ipValue.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;

                default:
                    return false;
            }
        }

        // The function below should be revised to allow a generic seial number
        //to be updated.
        // needs the ability to skip over certain characters used to
        //denote fields.
        private static string IncrementStr(string ipStr)
        {
            var myChars = ipStr.ToCharArray();

            var myIndex = myChars.Length - 1;
            var myInc = false;
            while (myIndex >= 0)
            {
                var myChar = IncrementChar(myChars[myIndex]);
                if (myChar < myChars[myIndex])
                {
                    myChars[myIndex] = myChar;
                    myIndex--;
                }
                else
                {
                    myChars[myIndex] = myChar;
                    myInc = true;
                    break;
                }
            }
            if (!myInc)
            {
                return "0" + new string(myChars);
            }
            else
            {
                return new string(myChars);
            }
        }

        private const string KeyChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVXYZabcdefghijklmnopqrstuvxyz";

        private static char IncrementChar(char ipChar)
        {
            do
            {
                if (ipChar == 'z')
                {
                    // Wrap around z to 0'
                    ipChar = '0';
                }
                else
                {
                    ipChar++;
                };
            } while (!KeyChars.Contains(ipChar.ToString()));
            return ipChar;
        }
    }
}