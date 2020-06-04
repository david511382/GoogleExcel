namespace GoogleExcel
{
    public class Scope
    {
        public static Scope SPREADSHEETS_READONLY => new Scope("https://www.googleapis.com/auth/spreadsheets.readonly");
        public static Scope SPREADSHEETS => new Scope("https://www.googleapis.com/auth/spreadsheets");
        public static Scope DRIVE_READONLY => new Scope("https://www.googleapis.com/auth/drive.readonly");
        public static Scope DRIVE_FILE => new Scope("https://www.googleapis.com/auth/drive.file");
        public static Scope DRIVE => new Scope("https://www.googleapis.com/auth/drive");

        private Scope(string value) { Value = value; }

        public string Value { get; }
    }
}
