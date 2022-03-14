using System;
using System.IO;
using CsvHelper;
using OfficeOpenXml;
using System.Globalization;
using CsvHelper.Configuration;
using System.Collections.Generic;


namespace Kpaul_challenge
{

    public abstract class Converter{
        public int rows;
        public int columns;
        //Every Converter will most likely use this function so we make an abstract version here so the children must implement it
        public abstract void convert(int row_per_file, string[] headers, string delimiter);

    }
    /**Abstract xlsx converter class so we can make children that implement the convert() function how they see fit with a few
    common required parameters and member variables**/
    public abstract class XlsxConverter: Converter{

        public ExcelWorksheet xl_sheet;

        public XlsxConverter(ExcelWorksheet worksheet, int row_count, int col_count){
            xl_sheet = worksheet;
            rows = row_count;
            columns = col_count;
        }
        //Xlsx converters must implement a check_cell method
        public abstract bool check_cell(object cell_content, int column_num);
    }
    public class XlsxToCsvConverter: XlsxConverter{

        public XlsxToCsvConverter(ExcelWorksheet worksheet, int row_count, int col_count): base(worksheet,row_count, col_count){}

        //Function that checks the current cell against the table requirements provided in the challenge
        public override bool check_cell(object cell_content, int column_num){
            if(cell_content == null && column_num != 8){
                return false;
            }
            else if(column_num == 8){
                if(cell_content != null){
                    return(cell_content.ToString().Length <= 12);
                } else {
                    return true;
                }
            }
            else if(column_num == 1 || column_num == 5){
                return (cell_content is double && (cell_content.ToString().Length <= 20));
            }
            else if (2 <= column_num && column_num <= 4 ){
                return (cell_content is string && (cell_content.ToString().Length <= 50 ));
            } 
            else if (column_num == 6 || column_num == 9){
                return(cell_content is string && (cell_content.ToString().Length <= 2));
            }
            else if(column_num == 7){
                return (cell_content is string && (cell_content.ToString().Length <= 300));
            } else {
                return false;
            }
        }

        /**
        *   Converts an .xlsx file into a csv given the desired amount of rows per file, the column headers, and 
        *   the desired delimiter to be used. The main nested loop starts by creating a file and a csv writer.
        *   we then start looping through every row and then every cell in each row and checking if they are within
        *   the given parameters. If we find an important missing item then we add it to the Error.xlsx.
        *   
        *   When the max number of rows per file is reached and we still have rows leftover, we will break the inner loop to create a new
        *   file and continue this whole process until we have reached the end of the original table. We then should have some .csv files saved
        *   in our current directory.
        *
        *   @param[in]      row_per_file        Max number of rows allowed per file
        *   @param[in]      headers             An array of strings containing the desired headers for the resulting .csv
        *   @param[in]      delimiter           The desired string to use as a delimiter
        **/
        public override void convert(int row_per_file, string[] headers, string delimiter){
            var config = new CsvConfiguration(CultureInfo.InvariantCulture){ Delimiter = delimiter, };
            int checkpoint = 1;
            int lines_written = 0;
            int file_num = 1;
            int err_row = 1;
            int err_col = 1;
            object content = null;
            List<object> current_row = new List<object>();

            bool missing_field = false;
            
            ExcelPackage error = new ExcelPackage();
            ExcelWorksheet error_sheet = error.Workbook.Worksheets.Add("Sheet1");

            foreach (string s in headers)//O(M)
            {
                error_sheet.Cells[err_row, err_col].Value = s;
                err_col++;
            }
            err_col = 1;

            /** N = num of files
            *   M = columns
            *   P = rows
            *   
            **/
            //Loop until we reach the end of the original table
            while(checkpoint <= rows){ //O(N)

                string path = @"file"+file_num+".csv";
                
                Console.WriteLine("FILE {0}", file_num);
                
                using var writer = new StreamWriter(path);
                using var csv_writer = new CsvWriter(writer, config);

                foreach(string s in headers){//O(N*M)
                    csv_writer.WriteField(s);
                }
                csv_writer.NextRecord();
                

                for (int i = checkpoint; i <= rows; i++) //BigO(P)
                {   

                    checkpoint = i+1;

                    for (int j = 1; j <= columns; j++) //BigO(P*M)
                    {
                        content = xl_sheet.Cells[i, j].Value;
                        
                        if(!check_cell(content, j)){
                            missing_field = true;
                            continue;
                        } else if(content is double && j == 5){
                            content = ((double)content*1.20);
                        } else if(j == 6){
                            content = "TW";
                        } else {
                            missing_field = false;
                        }
                        current_row.Add(content);
                    }

                    if(!missing_field){
                        current_row.ForEach(csv_writer.WriteField);
                        csv_writer.NextRecord();
                        lines_written++;
                    } else {

                        foreach (object o in current_row) //O(P*M)
                        {
                            error_sheet.Cells[err_row, err_col].Value = o;
                            err_col++;
                        }
                        err_col = 1;
                        err_row++;
                    }

                    if((lines_written%row_per_file) == 0 && lines_written != 0){
                        Console.WriteLine("Lines Written: {0} to {1}" , lines_written, path);
                        lines_written = 0;
                        break;
                    }
                    current_row.Clear();
                }
                file_num++;
                csv_writer.Dispose();

            }
            FileStream error_file = File.Create(@"Error.xlsx");
            error_file.Close();
            File.WriteAllBytes("Error.xlsx", error.GetAsByteArray());
            error.Dispose();
        }



    }

    /**Main for the example program. Just creating files and variables that will be used for conversion
    **/ 
    public class SpreadsheetConverter{
        static void Main(string[] args){
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(@"ProductList.xlsx");

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

            int row_count = worksheet.Dimension.Rows;
            int column_count = worksheet.Dimension.Columns;
            Console.WriteLine("Rows: {0}", row_count);
            Console.WriteLine("Columns: {0}", column_count);

            string[] headers = {"PID","Product ID","Mfr Name","Mfr P/N","Price","COO","Short Description","UPC","UOM"};
            string delim = "^";
            var converter = new XlsxToCsvConverter(worksheet, row_count, column_count);
            converter.convert(10000, headers, delim);
        }
    }
    
}