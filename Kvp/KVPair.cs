using System.Collections.Generic;
using System.Collections;
using System.Runtime.InteropServices;

namespace VBALibrary
{
    [ComVisible(true)]
    public class KVPair : IKVPair
    {
        public object Key { get; set; }
        public object Value { get; set; }

        public KVPair(dynamic ipKey, dynamic ipValue)
        {
            Key = ipKey;
            Value = ipValue;
        }
    }
}