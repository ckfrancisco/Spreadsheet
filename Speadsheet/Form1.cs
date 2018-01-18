using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SpreadsheetEngine;

namespace Speadsheet
{
    public partial class Form1 : Form
    {
        Spreadsheet mSpreadSheet;

        //description: form constructor
        //parameters: 
        //return: 
        public Form1()
        {
            InitializeComponent();
            mSpreadSheet = new Spreadsheet();

            for (int i = 0; i < 26; i++)                                    //initialize column names
                dataGridView1.Columns.Add(((char)(i + 65)).ToString(), ((char)(i + 65)).ToString());

            dataGridView1.Rows.Add(50);

            for (int i = 0; i < 50; i++)                                    //initialize row names
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            mSpreadSheet.PropertyChanged += OnSpreadsheetPropertyChanged;   //subscribe spreadsheet property change to OnSpreadSheetPropertyChanged
        }

        //description: once a cell property has been changed this function is called to update the UI
        //parameters: object sender and property changed arguements
        //return: 
        private void OnSpreadsheetPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Cell updateCell = sender as Cell;

            switch(e.PropertyName)
            {
                case "Value":
                case "Text":                                                                    //if value or text change then update the UI cell value
                    dataGridView1[updateCell.ColumnIndex, updateCell.RowIndex].Value =
                       mSpreadSheet.GetCell(updateCell.ColumnIndex, updateCell.RowIndex).Value;
                    break;

                case "BGColor":                                                                 //if color change then update the UI color value
                    dataGridView1[updateCell.ColumnIndex, updateCell.RowIndex].Style.BackColor = 
                         Color.FromArgb((int)mSpreadSheet.GetCell(updateCell.ColumnIndex, updateCell.RowIndex).BGColor);
                    break;

                default:
                    break;
            }
        }

        //description: demo button that reflects a change in cell text in the cell value and thus the UI
        //parameters: object sender (a cell) and property changed arguements
        //return: 
        private void button1_Click(object sender, EventArgs e)
        {
            Random x = new Random();

            for (int i = 0; i < 48; i++)
                mSpreadSheet.GetCell(x.Next(0, 26), x.Next(0, 50)).Text = "l l l l l";      //randomize locations of "l l l l l"

            mSpreadSheet.GetCell(2, 0).Text = "Find the 1";                                 //set "Find the 1" to C1

            mSpreadSheet.GetCell(x.Next(0, 26), x.Next(0, 50)).Text = "11111";              //randomize location of "11111"

            for (int i = 0; i < 50; i++)
                mSpreadSheet.GetCell(1, i).Text = "This is cell B" + (i + 1).ToString();    //set B column cells

            for (int i = 0; i < 50; i++)
                mSpreadSheet.GetCell(0, i).Text = "=B" + i.ToString(); ;                    //set A column to adjacent B column value
        }

        //description: displays cells text while editing
        //parameters: object sender and property changed arguements
        //return: 
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataGridView1[e.ColumnIndex, e.RowIndex].Value =            //during edit update UI view to cell's text
                mSpreadSheet.GetCell(e.ColumnIndex, e.RowIndex).Text;
        }

        //description: assigns cell a new value if new value is not null, and returns displays to cell's value
        //parameters: object sender and property changed arguements
        //return: 
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell uiCell = dataGridView1[e.ColumnIndex, e.RowIndex];
            Cell logicCell = mSpreadSheet.GetCell(e.ColumnIndex, e.RowIndex);

            if (uiCell.Value == null)                       //if the value is null set the cell to empty
                uiCell.Value = "";

            if (uiCell.Value.ToString() == logicCell.Text)  //if the text remains the same do not update any values
            {
                uiCell.Value = logicCell.Value;
                return;
            }

            mSpreadSheet.PushUndo(logicCell, "Text");       //push command into history and reset redo stack
            redoToolStripMenuItem.Text = "Redo";            //reset redo string
            redoToolStripMenuItem.Enabled = false;          //disable redo button

            undoToolStripMenuItem.Text = "Undo Text";       //set undo button to undo text change
            undoToolStripMenuItem.Enabled = true;           //enable undo button

            if (uiCell.Value != null)
                logicCell.Text =  uiCell.Value.ToString();  //if UI cell value isn't null update logic cell calue
            else
                logicCell.Text = "";                        //if UI cell value is null update logic cell value to ""

            uiCell.Value = logicCell.Value;                 //once the edit is done update UI view to cell's value
        }
    
        //description: choose a color from the color dialog and update the color stored in the logic cells
        //parameter: object sender and property changed arguements
        //return: 
        private void chooseBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color color = colorDialog.Color;                   //open color dialog to select color

                if (dataGridView1.SelectedCells.Count == 1)
                {
                    DataGridViewCell uiCell = dataGridView1.SelectedCells[0];
                    Cell logicCell = mSpreadSheet.GetCell(uiCell.ColumnIndex, uiCell.RowIndex);

                    mSpreadSheet.PushUndo(logicCell, "BGColor");    //push command into history and reset redo stack
                    redoToolStripMenuItem.Text = "Redo";            //reset redo string
                    redoToolStripMenuItem.Enabled = false;          //disable redo button

                    undoToolStripMenuItem.Text = "Undo BGColor";    //set undo button to undo bgcolor change
                    undoToolStripMenuItem.Enabled = true;           //enable undo button

                    logicCell.BGColor = (uint)color.ToArgb();       //once the edit is done update UI view to cell's color
                }

                else
                {
                    List<Cell> logicCells = new List<Cell>();

                    foreach (DataGridViewTextBoxCell uiCell in dataGridView1.SelectedCells)         //add all selected cells to list of cells
                        logicCells.Add(mSpreadSheet.GetCell(uiCell.ColumnIndex, uiCell.RowIndex));

                    mSpreadSheet.PushUndo(logicCells, "BGColor");                                   //push command into history and reset redo stack
                    redoToolStripMenuItem.Text = "Redo";                                            //reset redo string
                    redoToolStripMenuItem.Enabled = false;                                          //disable redo button

                    undoToolStripMenuItem.Text = "Undo BGColor";                                    //set undo button to undo bgcolor change
                    undoToolStripMenuItem.Enabled = true;                                           //enable undo button

                    foreach (DataGridViewTextBoxCell uiCell in dataGridView1.SelectedCells)         //once the edit is done update UI to all cell's colors  
                        mSpreadSheet.GetCell(uiCell.ColumnIndex, uiCell.RowIndex).BGColor = (uint)color.ToArgb();
                }
            }
        }

        //description: undo previous command
        //parameter: object sender and property changed arguements
        //return: 
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mSpreadSheet.Undo();                                            //undo command
            redoToolStripMenuItem.Text = "Redo " + mSpreadSheet.PeekRedo(); //update redo button to property of next redo command
            redoToolStripMenuItem.Enabled = true;                           //enable redo button

            if (mSpreadSheet.SizeUndo() == 0)                               //if no commands left on undo stack
            {                                                               //  resest text and disable
                undoToolStripMenuItem.Text = "Undo";
                undoToolStripMenuItem.Enabled = false;
            }

            else                                                            //else update string with property of next undo command
            {
                undoToolStripMenuItem.Text = "Undo " + mSpreadSheet.PeekUndo();
            }
        }

        //description: redo previous undo command
        //parameter: object sender and property changed arguements
        //return: 
        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mSpreadSheet.Redo();                                            //redo command
            undoToolStripMenuItem.Text = "Undo " + mSpreadSheet.PeekUndo(); //update undo button to property of next undo command
            undoToolStripMenuItem.Enabled = true;                           //enable undo button

            if (mSpreadSheet.SizeRedo() == 0)                               //if no commands left on redo stack
            {                                                               //  resest text and disable
                redoToolStripMenuItem.Text = "Redo";
                redoToolStripMenuItem.Enabled = false;
            }

            else                                                            //else update string with property of next redo command
            {
                redoToolStripMenuItem.Text = "Redo " + mSpreadSheet.PeekRedo();
            }
        }

        //description: load file into spreadsheet
        //parameter: object sender and property changed arguements
        //return: 
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog.FileName == "")                                          //if no file then cancel
                    return;

                mSpreadSheet.FilePath = openFileDialog.FileName;                            //set file path to string returned by dialog

                FileStream file = new FileStream(mSpreadSheet.FilePath, FileMode.Open);     //open file stream with file path

                mSpreadSheet.LoadFromFile(file);                                            //load stream into spreadsheet

                file.Close();

                undoToolStripMenuItem.Text = "Undo";                                        //reset undo/redo buttons
                undoToolStripMenuItem.Enabled = false;

                redoToolStripMenuItem.Text = "Redo";
                redoToolStripMenuItem.Enabled = false;
            }
        }

        //description: save spreadsheet into current file or newly determined file if none
        //parameter: object sender and property changed arguements
        //return: 
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mSpreadSheet.FilePath == "")                                        //if the file path is empty, than use save as function
                saveAsToolStripMenuItem_Click(sender, e);

            FileStream file = new FileStream(mSpreadSheet.FilePath, FileMode.Open); //open file stream with file path

            mSpreadSheet.SaveToFile(file);                                          //save spreadsheet to stream

            file.Close();
        }

        //description: save spreadsheet into newly determined file
        //parameter: object sender and property changed arguements
        //return: 
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog.FileName == "")                                          //if no file then cancel
                    return;

                mSpreadSheet.FilePath = saveFileDialog.FileName;                            //set file path to string returned by dialog

                FileStream file = new FileStream(mSpreadSheet.FilePath, FileMode.Create);   //open file stream with file path

                mSpreadSheet.SaveToFile(file);                                              //save spreadsheet to stream

                file.Close();
            }
        }
    }
}