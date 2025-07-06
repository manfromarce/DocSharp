using System;
using System.Diagnostics;

namespace DocSharp.Renderer.Core
{
    [DebuggerDisplay("PageNumber: {_number}")]
    internal class PageNumber : IEquatable<PageNumber>, IComparable<PageNumber>
    {
        public static readonly PageNumber None = new PageNumber(0);
        public static readonly PageNumber First = new PageNumber(1);

        private readonly int _number;

        public PageNumber(int number)
        {
            _number = number;
        }

        public PageNumber Next() => new PageNumber(_number + 1);

        public bool Equals(PageNumber? other)
        {
            return other is object
                && other._number == _number;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(_number);
        }

        public override bool Equals(object? obj)
        {
            return this.Equals(obj as PageNumber);
        }

        public int CompareTo(PageNumber? other)
        {
            if (other is null)
            {
                return 1;
            }
            return _number - other._number;
        }

        public static bool operator==(PageNumber p1, PageNumber p2)
        {
            return  p1.Equals(p2);
        }

        public static bool operator !=(PageNumber p1, PageNumber p2)
        {
            return !(p1 == p2);
        }

        public static implicit operator int(PageNumber pageNumber) => pageNumber._number;
    }
}
