using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Collections;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.TableComponents;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents
{
    public class RowsCollection : EditAbleCollection<RowEx>
    {
        private readonly Dictionary<RowEx, A.TableRow> _innerSdkDic;

        public override void Remove(RowEx innerRow)
        {
            if (!_innerSdkDic.ContainsKey(innerRow))
            {
                throw new ArgumentNullException(nameof(innerRow));
            }

            _innerSdkDic[innerRow].Remove();
            CollectionItems.Remove(innerRow);
        }

        public void RemoveAt(int index)
        {
            if (index < 0 || index >= CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = CollectionItems[index];
            Remove(innerRow);
        }

        public RowsCollection(IEnumerable<A.TableRow> sdkTblRows, IShapeContext spContext)
        {
            var count = sdkTblRows.Count();
            CollectionItems = new List<RowEx>(count);
            _innerSdkDic = new Dictionary<RowEx, A.TableRow>(count);
            foreach (var sdkRow in sdkTblRows)
            {
                var innerRow = new RowEx(sdkRow, spContext);

                _innerSdkDic.Add(innerRow, sdkRow);
                CollectionItems.Add(innerRow);
            }
        }
    }
}