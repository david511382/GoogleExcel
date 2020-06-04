using Google.Apis.Sheets.v4.Data;
using System;

namespace GoogleExcel.Models
{
    public struct RangePosition
    {
        public Position LeftTop, RightBottom;
        public string TableName;

        public RangePosition(string tableName, Position position)
            : this(tableName, position, position)
        {
        }

        public RangePosition(string tableName, Position leftTopPosition, Position rightBottomPosition)
        {
            LeftTop = leftTopPosition;
            RightBottom = rightBottomPosition;
            TableName = tableName;
        }

        public string GetStr()
        {
            string leftTopStr = LeftTop.GetStr();
            string rightBottomStr = RightBottom.GetStr();
            string result = $"{TableName}!{leftTopStr}:{rightBottomStr}";
            return result;
        }

        public GridRange GetGridRange(int sheetID)
        {
            if (LeftTop.IsAllColumn == !RightBottom.IsAllColumn ||
                LeftTop.IsAllRow == !RightBottom.IsAllRow)
            {
                throw new Exception("選取範圍不合法");
            }

            GridRange result = new GridRange()
            {
                SheetId = sheetID,
                StartRowIndex = (LeftTop.IsAllRow) ? 0 : LeftTop.Row - 1,
                EndRowIndex = (RightBottom.IsAllRow) ? 0 : RightBottom.Row,

                StartColumnIndex = (LeftTop.IsAllColumn) ? 0 : LeftTop.Column - 1,
                EndColumnIndex = (RightBottom.IsAllColumn) ? 0 : RightBottom.Column
            };

            return result;
        }
    }

}
