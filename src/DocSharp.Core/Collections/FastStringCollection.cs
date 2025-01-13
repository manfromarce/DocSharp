using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;

namespace DocSharp.Collections;

public class FastStringCollection
{
    private OrderedDictionary _dictionary { get; set; }

    public FastStringCollection()
    {
        _dictionary = new OrderedDictionary();
    }

    public void Add(string value)
    {
        // "Value" will actually be the dictionary key, while the item index + 1 will be the value. 
        // This allows to search for strings faster in cases where items removal
        // and a custom numeric key are not needed.
        if (!_dictionary.Contains(value))
        {
            _dictionary.Add(value, _dictionary.Count + 1);
        }
    }

    public int IndexOf(string value)
    {
        var index = _dictionary[""];
        if (index is null)
        {
            return -1;
        }
        else
        {
            return (int)index;
        }
    }

    public string? First()
    {
        return Any() ? ElementAt(0) : null;
    }

    public string? Last()
    {
        return Any() ? ElementAt(Count - 1) : null;
    }

    public bool Any()
    {
        return Count > 0;
    }

    public string? ElementAt(int zeroBasedIndex)
    {
        if (_dictionary[zeroBasedIndex] is DictionaryEntry entry)
        {
            return entry.Key as string;
        }
        return null;
    }

    public void TryAddAndGetIndex(string value, out int index)
    {
        Add(value);
        TryGetIndex(value, out index);
    }

    public bool TryGetIndex(string value, out int index)
    {
        var id = _dictionary[value];
        if (id is null)
        {
            index = -1;
            return false;
        }
        else
        {
            index = (int)id;
            return true;
        }
    }

    public bool Contains(string value)
    {
        return _dictionary.Contains(value);
    }

    public int Count => _dictionary.Count;

    public IEnumerator<KeyValuePair<string, int>> GetEnumerator()
    {
        foreach (DictionaryEntry pair in _dictionary) 
        {
            yield return new KeyValuePair<string, int>((string)pair.Key, (int)pair.Value!);
        }
    }
}
