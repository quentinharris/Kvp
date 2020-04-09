using System.Collections.Generic;
using System.Collections;
using System.Runtime.InteropServices;

namespace VBALibrary
{
    [ComVisible(true)]
    // [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[ComDefaultInterface(typeof(IKVPair))]
    public class KVPair : IKVPair
    {
        public object Key { get; set; }
        public object Value { get; set; }

        public KVPair(dynamic ipKey, dynamic ipValue)
        {
            Key = ipKey;
            Value = ipValue;
        }

        //public KVPair(KeyValuePair<dynamic, dynamic> ipPair)
        //{
        //    Key = ipPair.Key;
        //    Value = ipPair.Value;
        //}

        //public KVPair(DictionaryEntry ipPair)
        //{
        //    Key = ipPair.Key;
        //    Value = ipPair.Value;
        //}
    }
}