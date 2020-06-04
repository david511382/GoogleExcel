using System.Collections.Generic;

namespace GoogleExcel.Models
{
    public struct UpdateRange
    {
        public RangePosition Range;
        public IList<IList<object>> Data;

        public UpdateRange(string tableName, Position position, object value)
        {
            Data = new List<IList<object>>()
                {
                    new List<object>()
                    {
                        value
                    }
                };
            Range = new RangePosition(tableName, position);
        }
    }
}
