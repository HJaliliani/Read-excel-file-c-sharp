using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Asposecell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
 
        //Read All Rows And Columns, row by row and column by column and show them on MessageBox By Aspose.cells component
        private void button2_Click(object sender, EventArgs e)
        {
              //Open your template file.
            Workbook workbook = new Workbook(@"d:\test.xlsx");
            //Determine Sheet that we want to work with
            Worksheet worksheet = workbook.Worksheets[0];
            //Cell object to access the cells of the sheet
            Cells cells = worksheet.Cells;

            string tmpvalue="";
            for (int rw = 0; rw <= cells.MaxRow; rw++)
            {
                for (int clmn = 0; clmn < cells.MaxColumn; clmn++)
                {
                    //Read the cells and prevent from null exception error
                    tmpvalue = cells[rw, clmn].Value == null ? string.Empty : cells[rw, clmn].Value.ToString();
                
                    //show all  cells data
                    //  MessageBox.Show("[" +rw +","+clmn+"]="+ tmpvalue);
                }
                
            }
            MessageBox.Show("Finish reading all of the excell sheet ceels");

            //Get the AA column index. 
            //int col = CellsHelper.ColumnNameToIndex("A");
           
            // Access the "A1" cell in the sheet.
            //   Cell cell = cells["A2"];
            // Input the "Hello" text into the "A2" cell
            //  cell.PutValue("Hello");

            // Save the Excel file.
            // workbook.Save(@"d:\" + "output.xlsx", SaveFormat.Xlsx);
        }
    }
    
}
