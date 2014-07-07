using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;


namespace TRXFileGenerator
{
    public class Program
    {
        public static void Main(IList<string> args)
        {
            string fileName;

            int aborted = 0, passed = 0, failed = 0, notexecuted = 0;
            if (args.Count < 1)
            {
                return;
            }

            try
            {

                // Construct DirectoryInfo for the folder path passed in as an argument

                var di = new DirectoryInfo(args[0]);

                Application oXL;
                Workbook oWB;
                Worksheet oSheet;

                // Get a refrence to Excel
                oXL = new Application();

                // Create a workbook and add sheet
                oWB = oXL.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                oSheet = (Worksheet)oWB.ActiveSheet;
                oSheet.Name = "trx";
                oXL.Visible = true;
                oXL.UserControl = true;

                // Write the column names to the work sheet
                oSheet.Cells[1, 1] = "Processed File Name";
                oSheet.Cells[1, 2] = "Test ID";
                oSheet.Cells[1, 3] = "Test Name";
                oSheet.Cells[1, 4] = "Test Outcome";

                var row = 2;

                // For each .trx file in the given folder process it
                foreach (var file in di.GetFiles("*.trx"))
                {
                    fileName = file.Name;

                    // Deserialize TestRunType object from the trx file
                    var fileStreamReader = new StreamReader(file.FullName);

                    var xmlSer = new XmlSerializer(typeof(TestRunType));

                    var testRunType = (TestRunType)xmlSer.Deserialize(fileStreamReader);

                    // Navigate to UnitTestResultType object and update the sheet with test result information
                    foreach (var itob1 in testRunType.Items)
                    {
                        var resultsType = itob1 as ResultsType;

                        if (resultsType != null)
                        {
                            foreach (var itob2 in resultsType.Items)
                            {
                                var unitTestResultType = itob2 as UnitTestResultType;

                                if (unitTestResultType != null)
                                {
                                    oSheet.Cells[row, 1] = fileName;
                                    oSheet.Cells[row, 2] = unitTestResultType.testId;
                                    oSheet.Cells[row, 3] = unitTestResultType.testName;
                                    oSheet.Cells[row, 4] = unitTestResultType.outcome;

                                    if (0 == unitTestResultType.outcome.CompareTo("Aborted"))
                                    {
                                        oSheet.Rows.get_Range("A" + row, "D" + row).Interior.Color =
                                            ColorTranslator.ToWin32(Color.Yellow);
                                        aborted++;
                                    }
                                    else if (0 == unitTestResultType.outcome.CompareTo("Passed"))
                                    {
                                        oSheet.Rows.get_Range("A" + row, "D" + row).Interior.Color =
                                            ColorTranslator.ToWin32(Color.Green);
                                        passed++;
                                    }
                                    else if (0 == unitTestResultType.outcome.CompareTo("Failed"))
                                    {
                                        oSheet.Rows.get_Range("A" + row, "D" + row).Interior.Color =
                                            ColorTranslator.ToWin32(Color.Red);
                                        failed++;
                                    }
                                    else if (0 == unitTestResultType.outcome.CompareTo("NotExecuted"))
                                    {
                                        oSheet.Rows.get_Range("A" + row, "D" + row).Interior.Color =
                                            ColorTranslator.ToWin32(Color.SlateGray);
                                        notexecuted++;
                                    }
                                    row++;
                                }
                            }
                        }
                    }
                }

                row += 2;

                // Add summmary
                oSheet.Cells[row++, 1] = "Testcases Passed = " + passed;
                oSheet.Cells[row++, 1] = "Testcases Failed = " + failed;
                oSheet.Cells[row++, 1] = "Testcases Aborted = " + aborted;
                oSheet.Cells[row++, 1] = "Testcases NotExecuted = " + notexecuted;

                // Autoformat the sheet
                oSheet.Rows.get_Range("A1", "D" + row)
                    .AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic1, false, false, false, true, false, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}