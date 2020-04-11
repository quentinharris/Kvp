using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

namespace VBALibrary
{
    //[Guid("434C844C-9FA2-4EC6-AB75-45D3013D75BE")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    //[ComDefaultInterface(typeof(IKvp))]
    //[ComSourceInterfaces(typeof(IKvp))] obsolete
    public partial class Kvp : IKvp, IEnumerable

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


        // ToDo: The function below could be revised to allow a generic seial number
        // to be updated.
        // i.e. needs the ability to ignore certain characters used to
        // denote fields.
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
