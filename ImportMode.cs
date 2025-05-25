namespace WindowsFormsApp
{
    public enum ImportMode
    {
        Append,         // 追加
        ClearAndImport, // 清空写入
        ErrorIfExists   // 报错不操作
    }

    public class ImportSettings
    {
        public string TableName { get; set; } = string.Empty;
        public ImportMode Mode { get; set; } = ImportMode.Append;
        public bool CreateTableIfNotExists { get; set; } = true;
        public bool TrimStrings { get; set; } = true;
        public bool SkipEmptyRows { get; set; } = true;
        public int BatchSize { get; set; } = 1000;
    }
} 