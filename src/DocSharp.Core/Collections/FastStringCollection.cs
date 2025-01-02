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
