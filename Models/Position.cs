using System;

namespace GoogleExcel.Models
{
    public struct Position
    {
        private static readonly int COLUMN_A_CODE;

        static Position()
        {
            COLUMN_A_CODE = 'A';
        }

        public int Column
        {
            get
            {
                return _column;
            }
            set
            {
                if (value <= 0)
                {
                    if (IsAllRow)
                        throw new Exception("暫時不支援全選");
                    _column = 0;
                }
                else
                    _column = value;
            }
        }

        public int Row
        {
            get
            {
                return _row;
            }
            set
            {
                if (value < 0)
                {
                    if (IsAllColumn)
                        throw new Exception("暫時不支援全選");
                    _row = 0;
                }
                else
                    _row = value;
            }
        }

        public bool IsAllColumn => _column == 0;
        public bool IsAllRow => _row == 0;

        private const int ALPHABET_COUNT = 26;
        private int _column, _row;

        public Position(int row, int col)
        {
            _row = row;
            _column = col;
        }

        public string GetStr()
        {
            string colStr = "";
            string rowStr = "";

            if (!IsAllRow)
            {
                rowStr = Row.ToString();
            }

            if (IsAllColumn)
                return rowStr;
            else
            {
                int col = Column;
                while (col > 0)
                {
                    col--;
                    int code = COLUMN_A_CODE + col % ALPHABET_COUNT;
                    colStr = (char)code + colStr;
                    col /= ALPHABET_COUNT;
                }
            }

            return colStr + rowStr;
        }
    }

}
