/*
Nathan Young
SID# 10932954
CPTS 321
HW 7
*/


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetEngine;
using CptS321;
using System.IO;

namespace Spreadsheet_NYoung
{
    public partial class Form1 : Form
    {
        Spreadsheet thesheet = null;

        // **************************************************************************************************
        // **************************** INITIALIZE FORM *****************************************************
        // **************************************************************************************************


        public Form1() // A spreadsheet 
        {

            InitializeComponent();
            char header; // title for each column
            for (int i = 65; i < 91; i++) // label columns
            {
                header = (char)i;
                dataGridView1.Columns.Add(header.ToString(), header.ToString());
            }
            for (int i = 1; i < 51; i++) // label rows
            {
                DataGridViewRow newrow = new DataGridViewRow();
                newrow.HeaderCell.Value = i.ToString(); // title for each row
                dataGridView1.Rows.Add(newrow);
            }
            thesheet = new Spreadsheet(26, 50);
            dataGridView1.RowHeadersWidth = 50;
            thesheet.CellPropertyChanged += Thesheet_CellPropertyChanged; // subscribe to spreadsheet class
            Redo.Enabled = false;
            Undo.Enabled = false;
        }



        private void Thesheet_CellPropertyChanged(object sender, PropertyChangedEventArgs e) // UI/logic communication hub
        {
            if (e.PropertyName == "Value")
            {
                SpreadsheetCell thecaller = sender as SpreadsheetCell;
                int column = thecaller.ColumnIndex;
                int row = thecaller.RowIndex;
                (dataGridView1[column, row]).Value = thecaller.Value;
            }

            if (e.PropertyName == "BGColor")
            {
                SpreadsheetCell thecaller = sender as SpreadsheetCell;
                int column = thecaller.ColumnIndex;
                int row = thecaller.RowIndex;
                (dataGridView1[column, row]).Style.BackColor = System.Drawing.Color.FromArgb((int)thecaller.BGColor);
            }
        }

        // useless object

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // ********************************************* the button from HW3 ***********************************************

        private void TheButton_Click(object sender, EventArgs e) // 
        {
            Random number = new Random();
            for (int i = 0; i < 50; i++) // random
            { thesheet.GetCell(number.Next(0, 25), number.Next(0, 49)).Text = "Hello?"; }

            for (int i = 0; i < 49; i++) // label B
            {
                thesheet.GetCell(1, i).Text = "This is cell B" + (i + 1);
                thesheet.GetCell(0, i).Text = "=B" + (i + 1);
            }
        }

        // ********************************************* TEXT EDIT **********************************************************

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e) // cell text begin edit
        {
            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null)
            {
                SpreadsheetCell placeholder = thesheet.GetCell(e.ColumnIndex, e.RowIndex);
                dataGridView1[e.ColumnIndex, e.RowIndex].Value = placeholder.Text;
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e) // cell text end edit
        {
            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null)
            {
                SpreadsheetCell thecell = thesheet.GetCell(e.ColumnIndex, e.RowIndex);
                thesheet.AddUndo(thecell, thecell.Text);                                 // log change in undo stack
                Undo.Enabled = true;                                                     // enable undo button
                thecell.Text = (string)dataGridView1[e.ColumnIndex, e.RowIndex].Value;   // update text
                thesheet.ClearRedoStack();                                              // clear redo stack
                this.Redo.Enabled = false;                                              // disable redo button
                this.Undo.Text = "Undo Text Change";
            }
        }

        // ******************************************* COLOR EDIT *********************************************************

        private void BGColor_Click(object sender, EventArgs e) // change in Background color
        {
            System.Windows.Forms.ColorDialog changecolor = new ColorDialog();

            if (changecolor.ShowDialog() == DialogResult.OK)                                        // opens dialog
            {
                SpreadsheetCell thecell = null;
                SpreadsheetCell[] thelistofcells = new SpreadsheetCell[dataGridView1.SelectedCells.Count]; // all changed cells
                int i = 0;
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)                      // log cells current color
                {
                    thecell = thesheet.GetCell(cell.ColumnIndex, cell.RowIndex);
                    thelistofcells[i++] = thecell;
                }
                thesheet.AddUndo(thelistofcells);                                        // log change in undo stack
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)                      // update cells
                {
                    thecell = thesheet.GetCell(cell.ColumnIndex, cell.RowIndex);
                    thecell.BGColor = (uint)changecolor.Color.ToArgb();
                }
                thesheet.ClearRedoStack();                                                          // clear redo stack
                this.Undo.Enabled = true;                                                           // enable undo button
                this.Redo.Enabled = false;                                                          // disable redo button
                this.Undo.Text = "Undo Background Color Change";
            }
        }

        //**************************************** UNDO/REDO *******************************************************

        private void Undo_Click(object sender, EventArgs e)             // undo button
        {
            Undo.Text = thesheet.PopUndoStack();                        // execute undo
            if (thesheet.UndoIsEmpty() == true)
            { this.Undo.Enabled = false; }                              // disable undo button if undo stack is empty
            this.Redo.Enabled = true;                                   // enable redo button
            Redo.Text = thesheet.PeekRedoStack();                       // update redo text

        }

        private void Redo_Click(object sender, EventArgs e)             // redo button
        {
            Redo.Text = thesheet.PopRedoStack();                        // execute redo
            if (thesheet.RedoIsEmpty() == true)
            { this.Redo.Enabled = false; }                              // disable redo button if redo stack is empty
            this.Undo.Enabled = true;                                   // enable undo button
            Undo.Text = thesheet.PeekUndoStack();                       // update undo text
        }



        //*************************************** SAVE/LOAD **********************************************************
        private void Save_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog saveit = new SaveFileDialog(); // Helper code found on MSDN and Stackoverflow
            saveit.Title = "Save File";
            saveit.Filter = "txt files (*.txt)|*.txt| xml files (*.xml)|*.xml| All files (*.*)|*.*";
            saveit.FilterIndex = 2; // set default to txt file
            saveit.RestoreDirectory = true;

            if (saveit.ShowDialog() == DialogResult.OK)
            {
                if ((saveit.FileName) != null)
                {
                    using (Stream thestream = saveit.OpenFile())
                    {
                    thesheet.SaveSpreadsheet(thestream);
                    thestream.Dispose();
                    }
                }
            }
        }

        private void Load_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openit = new OpenFileDialog();
            openit.Title = "Open File"; // dialog
            openit.Filter = "txt files (*.txt)|*.txt| xml files (*.xml)|*.xml"; // limit to text and xml files
            openit.FilterIndex = 2;
            openit.RestoreDirectory = true; // return to original directory after close ... i think
            Undo.Enabled = false;
            Redo.Enabled = false;

            if (openit.ShowDialog() == DialogResult.OK) // OK button is clicked in dialog
            {
                using (Stream thestream = openit.OpenFile())
                {
                    thesheet.LoadSpreadsheet(thestream);
                    thestream.Dispose();
                }
            }
        }
    }
}
