using System;
using System.Diagnostics;

namespace DocSharp.Renderer.Models.Common
{
    [DebuggerDisplay("{_column}")]
    internal class PageColumn : IEquatable<PageColumn>, IComparable<PageColumn>
    {
        public static readonly PageColumn None = new PageColumn(0);
        public static readonly PageColumn First = new PageColumn(1);
        public static readonly PageColumn One = new PageColumn(1);

        private readonly int _column;

        public PageColumn(int column)
        {
            _column = column;
        }

        public PageColumn Next()
            => new PageColumn(_column + 1);

        public int CompareTo(PageColumn other)
        {
            return _column - other._column;
        }

        public bool Equals(PageColumn other)
        {
            return other is object
                && _column == other._column;
        }

        public override int GetHashCode()
        {
            return _column.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            return this.Equals(obj as PageColumn);
        }

        public static bool operator ==(PageColumn p1, PageColumn p2)
        {
            return p1.Equals(p2);
        }

        public static bool operator !=(PageColumn p1, PageColumn p2)
        {
            return !(p1 == p2);
        }

        public static implicit operator int(PageColumn pageColumn) => pageColumn._column;
    }
}
