using System.IO;

namespace ConvertData.Infrastructure.Parsing;

internal enum ExcelFileFormat
{
    ZipXlsx,
    OleXls
}

internal static class ExcelFileSignature
{
    public static ExcelFileFormat Detect(string path)
    {
        byte[] header = new byte[8];
        using (var fsSig = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            int read = fsSig.Read(header, 0, header.Length);
            if (read < 2)
                throw new InvalidDataException("File is too small");
        }

        bool isZip = header[0] == (byte)'P' && header[1] == (byte)'K';
        if (isZip)
            return ExcelFileFormat.ZipXlsx;

        bool isOle = header.Length >= 8
            && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0
            && header[4] == 0xA1 && header[5] == 0xB1 && header[6] == 0x1A && header[7] == 0xE1;

        if (isOle)
            return ExcelFileFormat.OleXls;

        throw new InvalidDataException("Unknown Excel format (not zip/xlsx and not OLE/xls)");
    }
}
