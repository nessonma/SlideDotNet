using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideDotNet.Models.TableComponents
{
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
            var sdkTblRows = _sdkGrFrame.Descendants<A.Table>().First().Elements<A.TableRow>();

            return new RowsCollection(sdkTblRows, _spContext);
        }

        #endregion Private Methods
    }
}