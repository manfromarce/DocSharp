using System;
using System.Diagnostics;
using DocSharp.Renderer.Core;

namespace DocSharp.Renderer.Models.Common
{
    [DebuggerDisplay("PN: {PageNumber}, {Column}/{_totalColumns}")]
    internal class PagePosition : IEquatable<PagePosition>, IComparable<PagePosition>
    {
        public static readonly PagePosition None = new PagePosition(PageNumber.None, PageColumn.None, PageColumn.None);

        private readonly PageColumn _totalColumns;

        public PagePosition(PageNumber pageNumber) : this(pageNumber, PageColumn.First, PageColumn.One)
        {
        }

        public PagePosition(PageNumber pageNumber, PageColumn column, PageColumn totalColumns)
        {
            this.PageNumber = pageNumber;
            this.Column = column;
            _totalColumns = totalColumns;
        }

        public PageNumber PageNumber { get; }

        public PageColumn Column { get; }

        public int PageColumnIndex => this.Column - 1;

        public PagePosition Next()
        {
            if(_totalColumns <= 0)
            {
                throw new Exception("Total Columns is zero");
            }

            return this.Column == _totalColumns
                ? this.NewPage()
                : this.NextColumn();
        }

        public PagePosition NextPage(PageColumn column, PageColumn totalColumns)
            => new PagePosition(this.PageNumber.Next(), column, totalColumns);

        public PagePosition SamePage(PageColumn column, PageColumn totalColumns)
            => new PagePosition(this.PageNumber, column, totalColumns);

        private PagePosition NextColumn()
            => new PagePosition(this.PageNumber, this.Column.Next(), _totalColumns);

        private PagePosition NewPage()
            => new PagePosition(this.PageNumber.Next(), PageColumn.First, _totalColumns);

        public int CompareTo(PagePosition other)
        {
            var pnSign = this.PageNumber.CompareTo(other.PageNumber);
            if(pnSign != 0)
            {
                return pnSign;
            }

            return this.Column.CompareTo(other.Column);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(this.PageNumber.GetHashCode(), this.Column.GetHashCode());
        }

        public override bool Equals(object obj)
        {
            return this.Equals(obj as PagePosition);
        }

        public bool Equals(PagePosition other)
        {
            return other is object
                && other.PageNumber == this.PageNumber
                && other.Column == this.Column;
        }

        public static bool operator ==(PagePosition p1, PagePosition p2)
        {
            return p1.Equals(p2);
        }

        public static bool operator !=(PagePosition p1, PagePosition p2)
        {
            return !(p1 == p2);
        }
    }
}
