using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class CollectionExtensions
{
# if NETFRAMEWORK
    public static bool TryPeek<T>(this Stack<T> stack, out T? value)
    {
        if (stack.Count > 0)
        {
            value = stack.Peek();
            return true;
        }
        value = default;
        return false;
    }

    public static TValue GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
    {
        if (dictionary is null)
        {
            return defaultValue;
        }
        return dictionary.TryGetValue(key, out TValue? value) ? value : defaultValue;
    }
#endif

    public static Stack<T> Clone<T>(this Stack<T> original) 
    {
        var arr = new T[original.Count];
        original.CopyTo(arr, 0);
        Array.Reverse(arr);
        return new Stack<T>(arr);
    }

    public static T? TryPop<T>(this Stack<T> stack)
    {
        if (stack.Count > 0)
            return stack.Pop();
        return default;
    }

    public static T? TryPeek<T>(this Stack<T> stack)
    {
        if (stack.Count > 0)
            return stack.Peek();
        return default;
    }
}
