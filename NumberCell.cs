namespace FormExcel
{
    // struct for using index
    public struct TheIndex
    {
        public int row;
        public int column;
    }
    class NumberCell
    {
        const int NumberOfLetter = 26; //the number of latters in english
        const int TheFirstLetterInASCII = 65; //the first index of char type in the ASCII
        const int TheLastLetterInASCII = 90; //the last index of char type in the ASCII
        const int TheFirstNumberInASCII = 57; //the first index of int type in the ASCII
        const int TheLastNumberInASCII = 48; //the last index of int type in the ASCII

        public static string ToIndexSystem(int x)
        {
            x++;
            int mod;
            string ColumnName = "";

            if (x == 0) return ((char)(TheFirstLetterInASCII - 1)).ToString();

            while (x > 0)
            {
                mod = (x - 1) % NumberOfLetter;
                ColumnName = ((char)(TheFirstLetterInASCII + mod)).ToString() + ColumnName;
                x = (x - mod) / NumberOfLetter;
            }
            return ColumnName;
        }

        public static TheIndex FromIndexSystem(string index)
        {
            TheIndex ans = new TheIndex();
            ans.column = 0;
            ans.row = 0;

            for (int i = 0; i < index.Length; i++)
            {
                if (index[i] >= (TheFirstLetterInASCII - 1) && index[i] <= TheLastLetterInASCII)
                {
                    ans.column *= NumberOfLetter;
                    ans.column += (index[i]) - (TheFirstLetterInASCII - 1);
                }
                else if (index[i] >= TheLastNumberInASCII && index[i] <= TheFirstNumberInASCII)
                {
                    ans.row *= 10;
                    ans.row += index[i] - TheLastNumberInASCII;
                }
            }
            ans.column--;
            return ans;
        }
    }
}

