using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace Pings.Utils
{
    /// <summary>
    /// DataGridViewでのソートを可能にするために、BindingListを拡張したクラス
    /// </summary>
    public class SortableBindingList<T> : BindingList<T>
    {
        private List<T> originalList;
        private bool isSorted;
        private ListSortDirection sortDirection;
        private PropertyDescriptor sortProperty;

        public SortableBindingList(IList<T> list) : base(list)
        {
            originalList = (List<T>)list;
        }

        protected override bool SupportsSortingCore => true;
        protected override bool IsSortedCore => isSorted;
        protected override PropertyDescriptor SortPropertyCore => sortProperty;
        protected override ListSortDirection SortDirectionCore => sortDirection;

        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            originalList = (List<T>)this.Items;
            Type propType = prop.PropertyType;

            Comparison<T> comparer = (T x, T y) =>
            {
                object xValue = prop.GetValue(x);
                object yValue = prop.GetValue(y);

                if (xValue == null && yValue == null) return 0;
                if (xValue == null) return (direction == ListSortDirection.Ascending) ? -1 : 1;
                if (yValue == null) return (direction == ListSortDirection.Ascending) ? 1 : -1;

                if (xValue is IComparable comparableX)
                {
                    return comparableX.CompareTo(yValue) * (direction == ListSortDirection.Ascending ? 1 : -1);
                }
                return 0;
            };

            originalList.Sort(comparer);
            isSorted = true;
            sortProperty = prop;
            sortDirection = direction;
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        protected override void RemoveSortCore()
        {
            isSorted = false;
            sortDirection = ListSortDirection.Ascending;
            sortProperty = null;
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        public void Sort(PropertyDescriptor prop, ListSortDirection direction)
        {
            ApplySortCore(prop, direction);
        }
    }
}