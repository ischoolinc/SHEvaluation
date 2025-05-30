﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 班級定期評量成績單_固定排名
{
    public static class Tool
    {
     public static TValue GetValueOrDefault<TKey, TValue>
     (this IDictionary<TKey, TValue> dictionary,
      TKey key,
      TValue defaultValue)
        {
            TValue value;
            return dictionary.TryGetValue(key, out value) ? value : defaultValue;
        }

    }
}
