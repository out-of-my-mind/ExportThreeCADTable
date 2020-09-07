using System.Collections.Generic;
using System.Linq;
using System.Collections;

namespace Util
{
    /// <summary>
    /// 拓展方法类
    /// </summary>
    public static class ExtensionMethod
    {
        //keyValue类型  <key,IList>
        public static Dictionary<TKey, TValue> add<TKey, TValue, T>(this Dictionary<TKey, TValue> source, TKey key, T value) where TValue : IList, new()
        {
            if (source.Keys.Contains(key))
            {
                source[key].Add(value);
            }
            else
            {
                source.Add(key, new TValue() { value });
            }
            return source;
        }
        /// <summary>如果已经存在key数据，添加的时候是否覆盖
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="source"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="isCover"></param>
        /// <returns></returns>
        public static void addIsCover<TKey, TValue>(this Dictionary<TKey, TValue> source, TKey key, TValue value, bool isCover)
        {
            if (source.Keys.Contains(key))
            {
                if (isCover)
                {
                    source[key] = value;
                }
            }
            else
            {
                source.Add(key, value);
            }
            //return source;
        }
    }
}

