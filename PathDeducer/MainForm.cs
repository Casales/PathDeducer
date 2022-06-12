using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PathDeducer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button_SelectExcel_Click(object sender, EventArgs e)
        {
            

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                openExcel(openFileDialog.FileName);
            }

            
        }

        private void openExcel(string path)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Add(path);
                oSheet = oWB.ActiveSheet;
                oXL.Visible = true;

                var emptyIndex = getEmptyCellIndex(oSheet);
                var emptyRowIndex = getEmptyRowIndex(oSheet);
                var storyIndex = getPathCellIndex(oSheet);

                if(storyIndex != -1 && emptyIndex != -1)
                {
                    oSheet.Cells[1, emptyIndex] = "StoryPathID";
                    for (int i = 2; i < emptyRowIndex; i++)
                    {
                        var storyPath = (string)oSheet.Cells[i, storyIndex].Value;
                        if (storyPath == null) continue;
                        var pathID = getPathIdentifier(storyPath);
                        oSheet.Cells[i, emptyIndex] = pathID;
                    }
                }

                Debug.WriteLine($"Empty: {emptyIndex}");
                Debug.WriteLine($"Story: {storyIndex}");
                Debug.WriteLine($"Rows: {emptyRowIndex}");
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private int getEmptyCellIndex(Microsoft.Office.Interop.Excel._Worksheet oSheet)
        {
            for (int index = 1; index < oSheet.Cells.Columns.Count; index++)
            {
                dynamic cellContent = oSheet.Cells[1, index].Value;
                if (cellContent == null) return index;
            }

            return -1;
        }

        private int getEmptyRowIndex(Microsoft.Office.Interop.Excel._Worksheet oSheet)
        {
            for (int index = 1; index < oSheet.Cells.Rows.Count; index++)
            {
                dynamic cellContent = oSheet.Cells[index, 1].Value;
                if (cellContent == null) return index;
            }

            return -1;
        }

        private int getPathCellIndex(Microsoft.Office.Interop.Excel._Worksheet oSheet)
        {
            for(int index = 1; index < oSheet.Cells.Columns.Count; index++)
            {
                var cellContent = (string)oSheet.Cells[1, index].Value;
                if (cellContent.Contains("storyPath")) return index;
            }

            return -1;
        }

        private int getPathIdentifier(string path)
        {
            int pathID = 0;
            if (path.Contains("1b-Konsequenzen")) pathID += 1;
            if (path.Contains("2b-Konsequenzen")) pathID += 2;
            if (path.Contains("3b-Konsequenzen")) pathID += 4;
            if (path.Contains("4b-Konsequenzen")) pathID += 8;

            return pathID;
        }
    }
}
