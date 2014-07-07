using System;
using System.IO;

internal class LogFile
{
    private const string LogLocation = "Result.txt";

    public LogFile()
    {
    }

    public void ResultLog(String csvURL, String browserURL, String expectedURL, String textResult)
    {
        var outtxt = new FileInfo(LogLocation);
        var logline = outtxt.AppendText();
        logline.WriteLine("{0},{1},{2},{3}", csvURL, browserURL, expectedURL, textResult);

        // flush and close file.
        logline.Flush();
        logline.Close();
    }
}