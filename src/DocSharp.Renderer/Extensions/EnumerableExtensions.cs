using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace DocSharp.Renderer
{
    internal static class EnumerableExtensions
    {
        public static T FindPreviousOrDefault<T>(this IEnumerable<T> source, T item) where T : class
        {
            return source
                .Reverse()
                .SkipWhile(i => i != item)
                .Skip(1)
                .FirstOrDefault();
        }

        public static T FindNextOrDefault<T>(this IEnumerable<T> source, T item) where T : class
        {
            return source
                .SkipWhile(i => i != item)
                .Skip(1)
                .FirstOrDefault();
        }

        public static V MaxOrDefault<T, V>(this ICollection<T> source, Func<T, V> selector, V ifEmptyValue)
        {
            return source.Count == 0
                ? ifEmptyValue
                : source.Max(selector);
        }

        public static IEnumerable<(T value, int index)> SelectWithIndeces<T>(this IEnumerable<T> source, params int[] indeces)
        {
            var temp = source.ToArray();
            return indeces.Select(i => (temp[i], i));
        }

        public static IEnumerable<T> SelectByIndeces<T>(this IEnumerable<T> source, params int[] indeces)
        {
            return source
                .Where((item, index) => indeces.Contains(index));
        }

        public static IEnumerable<T> Update<T>(this IEnumerable<T> source, params (T value, int index)[] values)
        {
            var copy = source.ToArray();
            foreach(var v in values)
            {
                copy[v.index] = v.value;
            }
            return copy;
        }

        public static IEnumerable<T> MergeAndFilter<T>(T value, IEnumerable<T> otherValues, Func<T, bool> predicate)
        {
            return new[] { value }
                .Concat(otherValues)
                .Where(v => predicate(v));
        }

        public static Stack<T> ToStack<T>(this IEnumerable<T> source, bool reverseOrder = true)
        {
            var content = reverseOrder
                ? source.Reverse()
                : source;

            return new Stack<T>(content);
        }

        public static void Push<T>(this Stack<T> stack, IEnumerable<T> items, bool reverseOrder = true)
        {
            var order = reverseOrder
                ? items.Reverse()
                : items;

            foreach (var item in order)
            {
                stack.Push(item);
            }
        }

        public static int IndexOf<T>(this IEnumerable<T> source, Func<T, bool> predicate)
        {
            var i = 0;
            foreach(var e in source)
            {
                if (predicate(e))
                {
                    return i;
                }

                i++;
            }

            return -1;
        }
    }
}
