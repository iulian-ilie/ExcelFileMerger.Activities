using System;
using System.Collections.Generic;
using System.Linq;
using System.Activities;
using System.ComponentModel;
using System.IO;
using OfficeOpenXml;

namespace ExcelFileMerger
{
    public class ExcelFileMerger : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("An array of .XLSX file paths.\nThe first one will be used as a template.")]
        public InArgument<string[]> InputFiles { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("The .XLSX file path where the output is saved.\nIf the file exists it will be replaced.")]
        public InArgument<string> OutputFile { get; set; }

        [Category("Options")]
        [Description("Keep headers for all files.\nTemplate headers are always kept.")]
        public bool KeepHeaders { get; set; }

        [Category("Options")]
        [Description("Merge files even if they have a different number of columns.\nIf this is not checked the files will be skipped.")]
        public bool IgnoreColumnDifferences { get; set; }

        [Category("Options")]
        [Description("Adds a column with the path of the source file.")]
        public bool AddFileNames { get; set; }

        [Category("Output")]
        [Description("An array with the files that have been merged succesfuly.")]
        public OutArgument<string[]> FilesMerged { get; set; }

        [Category("Output")]
        [Description("An array with the files that have been skipped.")]
        public OutArgument<string[]> FilesSkiped { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var inputFiles = InputFiles.Get(context);
            var outputFile = OutputFile.Get(context);

            var mergeFiles = MergeFiles(inputFiles, outputFile, KeepHeaders, IgnoreColumnDifferences, AddFileNames);

            FilesMerged.Set(context, mergeFiles.FilesMerged.ToArray());
            FilesSkiped.Set(context, mergeFiles.FilesSkiped.ToArray());
        }

        private ResultSet MergeFiles(string[] inputFiles, string outputFile, bool keepHeaders, bool ignoreColumnDifferences, bool addFileNames)
        {
            int headerOffset = 0;
            int merged = 0;
            int skipped = 0;
            List<string> filesMerged = new List<string>();
            List<string> filesSkipped = new List<string>();

            try
            {
                var saveFile = new FileInfo(outputFile);

                // check if at least 2 files are supplied
                if (inputFiles.Count() <= 1)
                {
                    throw new ArgumentException("There must be at least 2 input files.");
                }

                //check if files exist and are .XLSX
                var filesToMerge = new List<string>();
                foreach (var file in inputFiles)
                {
                    if (!File.Exists(file))
                    {
                        throw new FileNotFoundException(string.Format("File not found: '{0}'",file));
                    }
                    else if(Path.GetExtension(file) != ".xlsx")
                    {
                        throw new ArgumentException(string.Format("The file '{0}' is invalid. Can only merge .XLSX files.",file));
                    }
                    else
                    {
                        filesToMerge.Add(file);
                    }
                }

                if (!keepHeaders) { headerOffset++; }
                var fileTemplate = new FileStream(filesToMerge[0], FileMode.Open, FileAccess.Read);

                using (ExcelPackage templateWorkBook = new ExcelPackage(fileTemplate))
                {
                    // the first file in the array is used as a template
                    ExcelWorksheet wsTemplate = templateWorkBook.Workbook.Worksheets[1];
                    var referenceColumn = wsTemplate.Dimension.End.Column;

                    // if addFileNames is true a new column in added with the source file
                    if (addFileNames)
                    {
                        for (int z = 2; z < wsTemplate.Dimension.End.Row + 1; z++)
                        {
                            wsTemplate.Cells[z, referenceColumn + 1].Value = filesToMerge[0];
                        }
                    }

                    // loops through all the files and adds them to the template
                    for (int i = 1; i < filesToMerge.Count; i++)
                    {
                        // gets the ranges for available cells in the template workbook
                        var maxRowTemplate = wsTemplate.Dimension.End.Row;
                        var maxColumnTemplate = wsTemplate.Dimension.End.Column;

                        var mergedFile = new FileStream(filesToMerge[i], FileMode.Open, FileAccess.Read);
                        using (ExcelPackage excelWorkbook = new ExcelPackage(mergedFile))
                        {
                            ExcelWorksheet ws = excelWorkbook.Workbook.Worksheets[1];
                            // gets the ranges for available cells in the workbook
                            var maxRow = ws.Dimension.End.Row;
                            var maxColumn = ws.Dimension.End.Column;

                            if (maxColumn == referenceColumn || ignoreColumnDifferences)
                            {
                                for (int x = 0 + headerOffset; x < maxRow; x++)
                                {
                                    for (int y = 0; y < maxColumn; y++)
                                    {
                                        wsTemplate.SetValue(maxRowTemplate + x + 1 - headerOffset, y + 1, ws.Cells[x + 1, y + 1].Value);
                                    }
                                    if (addFileNames)
                                    {
                                        wsTemplate.SetValue(maxRowTemplate + x + 1 - headerOffset, maxColumn + 1, filesToMerge[i]);
                                    }
                                }
                                merged++;
                                filesMerged.Add(filesToMerge[i]);
                            }
                            else
                            {
                                skipped++;
                                filesSkipped.Add(filesToMerge[i]);
                            }
                        }
                    }

                    // save the output file
                    templateWorkBook.SaveAs(saveFile);
                }

            }
            catch (Exception)
            {
                throw;
            }

            return new ResultSet
            {
                FilesMerged = filesMerged,
                FilesSkiped = filesSkipped
            };
        }
    }

    public class ResultSet
    {
        public List<string> FilesMerged { get; set; }
        public List<string> FilesSkiped { get; set; }
    }
}