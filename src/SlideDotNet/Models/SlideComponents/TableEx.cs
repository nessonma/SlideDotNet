using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Collections;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.TableComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideXML.Models.SlideComponents
{
    public class RowsCollection : EditAbleCollection<RowEx>
    {
        public override void Remove(RowEx item)
        {
            throw new NotImplementedException();
        }

        public void RemoveAt(int index)
        {

        }

        public RowsCollection(IEnumerable<RowEx> rows)
        {
            CollectionItems = new List<RowEx>(rows);
        }
    }

    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableEx
    {
        #region Fields

        private Lazy<RowsCollection> _rowsCollection;
        private readonly P.GraphicFrame _sdkGrFrame;
        private readonly IShapeContext _spContext;

        #endregion Fields

        #region Properties

        public RowsCollection Rows => _rowsCollection.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TableEx"/> class.
        /// </summary>
        public TableEx(P.GraphicFrame xmlGrFrame, IShapeContext spContext)
        {
            _sdkGrFrame = xmlGrFrame ?? throw new ArgumentNullException(nameof(xmlGrFrame));
            _spContext = spContext ?? throw new ArgumentNullException(nameof(spContext));
            _rowsCollection = new Lazy<RowsCollection>(GetRowsCollection());
        }

        #endregion Constructors

        #region Private Methods

        private RowsCollection GetRowsCollection()
        {
            var sdkTableRows = _sdkGrFrame.Descendants<A.Table>().First().Elements<A.TableRow>();
            var rows = new List<RowEx>(sdkTableRows.Count());
            foreach (var sdkTblRow in sdkTableRows)
            {
                rows.Add(new RowEx(sdkTblRow, _spContext));
            }

            return new RowsCollection(rows);
        }

        #endregion Private Methods
    }
}