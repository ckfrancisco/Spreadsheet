using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Xml.Linq;
using System.Collections;
using System.Text.RegularExpressions;

namespace SpreadsheetEngine
{
    public class Spreadsheet
    {
        int mColumns;
        int mRows;
        Cell[,] mCells;
        History mHistory;
        string mFilePath;

        Dictionary<Cell, HashSet<Cell>> mCellDependencies;

        public event PropertyChangedEventHandler PropertyChanged;

        //description: signals property changed
        //parameters: object sender (a cell) and property changed arguements
        //return: 
        protected void OnCellPropertyChanged(object sender, string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(sender, new PropertyChangedEventArgs(propertyName));    //pass cell as object sender
            }
        }

        //description: signals property changed and the cell(sender)'s text to the updated string or variable value
        //parameters: object sender and property changed arguements
        //return: 
        private void CellPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            CellValue updateCell = sender as CellValue;

            switch (e.PropertyName)                                                             //use event argument to determine the next event
            {
                case "Text":

                    UpdateCell(updateCell);                                                     //signal the change in text to update the value in the logic layer

                    goto case "Value";                                                          //continue to update the value in the UI layer

                case "Value":

                    OnCellPropertyChanged(updateCell, e.PropertyName);                          //signal change to UI for an update

                    if ((updateCell as Cell).Value == "!(circular-reference)")                  //if circular-reference end updates
                        return;

                    if ((updateCell as Cell).Value == "!(self-reference)")                      //if self-reference end updates
                        return;

                    foreach (Cell depCell in mCellDependencies[updateCell])
                        CellPropertyChanged(depCell, new PropertyChangedEventArgs("Value"));    //signal change to UI for all cells depending on the updating cell
                    break;

                case "BGColor":
                    OnCellPropertyChanged(updateCell, e.PropertyName);                          //signal change to UI for all cells depending on the updating cell
                    break;

                default:
                    break;
            }
        }

        class CellValue : Cell
        {
            //description: cellvalue constructor
            //parameters: column index, row index, text, and value
            //return: 
            public CellValue(int newColumnIndex = 0, int newRowIndex = 0, string newText = "", string newValue = "")
                : base(newColumnIndex, newRowIndex, newText, newValue)
                {}

            //description: value property setter
            //parameters: value
            //return:
            public new string Value { set { mValue = value; } }
        }

        //description: spreadsheet constructor
        //parameters: amount of columns and rows
        //return: 
        public Spreadsheet(int newColumns = 26, int newRows = 50)
        {
            mColumns = newColumns;
            mRows = newRows;
            mCells = new Cell[newColumns, newRows];
            mCellDependencies = new Dictionary<Cell, HashSet<Cell>>();
            mHistory = new History();
            mFilePath = "";


            for (int x = 0; x < mColumns; x++)
            {
               for(int y = 0; y < mRows; y++)
                {
                    mCells[x, y] = new CellValue(x, y);
                    mCells[x, y].PropertyChanged += CellPropertyChanged;    //subscribe each cell property change to CellPropertyChanged

                    mCellDependencies.Add(GetCell(x, y), new HashSet<Cell>());
                }
            }
        }

        //description: return cell within spreadsheet indicated by column and row
        //parameters: column and row
        //return: cell or null
        public Cell GetCell(int column, int row)
        {
            try                 //try to access and return element indicated by columns and row
            {
                return mCells[column, row];
            }
            catch(IndexOutOfRangeException)
            {
                return null;    //if out of bounds return null
            }
        }

        //description: column count property getter
        //parameters: 
        //return: column count
        public int ColumnCount { get { return mColumns; } }

        //description: row count property getter
        //parameters: 
        //return: row count
        public int RowCount { get { return mRows; } }

        //description: file path property getter and setter
        //parameters: 
        //return: file path
        public string FilePath { get { return mFilePath; } set { mFilePath = value; } }

        #region RegExp Variable Helper Functions
        //description: parse the expression string to return a hashet of the variable names
        //parameter: expression string
        //return: hashset of variable names
        //contributors: Peter Qafoku
        public static HashSet<string> FindVars(string exp)
        {
            HashSet<string> vars = new HashSet<string>(Regex.Matches(exp, @"[A-Za-z]+\d+").Cast<Match>().Select(matches => matches.Value));

            return vars;
        }

        //description: parse the expression string to return a hashet of the variable names
        //parameter: expression string
        //return: hashset of variable names
        public static HashSet<string> CheckVars(string exp)
        {
            HashSet<string> vars = new HashSet<string>(Regex.Matches(exp, @"[A-Za-z]+\d+|\d+").Cast<Match>().Select(matches => matches.Value));
            HashSet<string> tokens = new HashSet<string>(exp.Split('+', '-', '*', '/'));

            if (vars.Count != tokens.Count) //compares numbers of variables/numbers with number of tokens
                return null;                // if they don't match then an unknown name has been used

            else                            //else return hash of all variables
                return new HashSet<string>(Regex.Matches(exp, @"[A-Za-z]+\d+").Cast<Match>().Select(matches => matches.Value));
        }
        #endregion

        //description: update cell dependencies and values
        //parameters: the cell being updated
        //return: 
        private void UpdateCell(Cell updateCell)
        {
            HashSet<Cell> visCells = new HashSet<Cell>();

            UpdateTableAdd(updateCell);                             //add new dependencies
            UpdateTableDelete(updateCell);                          //remove old dependencies
            UpdateCellValue(updateCell as CellValue, visCells);     //update cell value
        }

        //description: add updating cell to new dependent cell's hashset of dependencies
        //parameters: the cell being updated
        //return: 
        private void UpdateTableAdd(Cell updateCell)
        {
            HashSet<string> prevVars = new HashSet<string>();
            HashSet<string> newVars = new HashSet<string>();

            if(updateCell.mTree != null)
                prevVars = updateCell.mTree.GetVars();          //create hashset of variables from previous expression tree

            if(updateCell.Text != "")
                newVars = FindVars(updateCell.Text);            //create hashet of variable parsed from new text

            foreach (string var in newVars)                     //for each new dependent cell add the updating cell
            {                                                   //  to the dependent cell's list of dependencies
                if (!prevVars.Contains(var))
                {
                    Cell depCell = GetCell(char.ToUpper(var[0]) - 65, int.Parse(var.Substring(1)) - 1);

                    if(depCell != null)
                        mCellDependencies[depCell].Add(updateCell);
                }
            }
        }

        //description: remove updating cell from old dependent cell's hashset of dependencies
        //parameters: the cell being updated
        //return: 
        private void UpdateTableDelete(Cell updateCell)
        {
            HashSet<string> prevVars = new HashSet<string>();
            HashSet<string> newVars = new HashSet<string>();

            if (updateCell.mTree != null)
                prevVars = updateCell.mTree.GetVars();          //create hashset of variables from previous expression tree

            if (updateCell.Text != "")
                newVars = FindVars(updateCell.Text);            //create hashet of variable parsed from new text

            foreach (string var in prevVars)                    //for each dependent cell no longer required remove the updating cell
            {                                                   //  from the previously dependent cell's dependencies
                if (!newVars.Contains(var))
                {
                    Cell depCell = GetCell(char.ToUpper(var[0]) - 65, int.Parse(var.Substring(1)) - 1);

                    if (depCell != null)
                        mCellDependencies[depCell].Remove(updateCell);
                }
            }
        }

        //description: update a cell's value after dependencies have been redetermined
        //parameters: the cell whose value must be updated
        //return: 
        private void UpdateCellValue(CellValue updateCell, HashSet<Cell>visCells)
        {
            if (visCells.Contains(updateCell))                              //if circular-references then return value and end updates
            {
                updateCell.Value = "!(circular-reference)";
                return;
            }

            else
                visCells.Add(updateCell);

            if (updateCell.Text.Length > 1 && updateCell.Text[0] == '=')    //if the string is larger than one char and starts with '='
            {
                updateCell.Value = EvalTree(updateCell);                    //  determine new value via expression tree

                if ((updateCell as Cell).Value == "!(self-reference)")      //if self-reference end updates
                    return;
            }

            else
            {
                updateCell.mTree = null;                                    //else reset cell tree to remove previous variables
                updateCell.Value = updateCell.Text;                         //set string to text
            }

            foreach (Cell depCell in mCellDependencies[updateCell])
            {
                UpdateCellValue(depCell as CellValue, visCells);            //update the value of all cells who depend on the updating cell
            }
        }

        //description: determines value from expression tree evaluation
        //parameters: the cell whose value must be updated
        //return: value of expression converted to a string
        private string EvalTree(Cell cell)
        {
            HashSet<string> vars = CheckVars(cell.Text.Substring(1).ToUpper());     //retrieve values declared in expression

            cell.mTree = new ExpTree(cell.Text.Substring(1).ToUpper());             //build the tree from the expression
                                                                                    //  IMPORTANT: tree must be built before checks to maintain variables

            if (vars == null)                                                       //if vars is null then a non-cell variable error was found
                return "!(name)";

            if (vars.Contains(string.Format("{0}{1}", (char)(cell.ColumnIndex + 65), cell.RowIndex + 1)))
                return "!(self-reference)";                                         //if the vars contains the cell then a self-reference error was found

            foreach (string var in vars)
            {
                double value = 0;                                                   //default variable value

                Cell variableCell = GetCell(var[0] - 65, int.Parse(var.Substring(1)) - 1);

                if (variableCell == null)                                           //if cell is null then out of bounds error was found
                    return "!(bounds)";

                if (variableCell.Value != "")
                {
                    if (double.TryParse(variableCell.Value, out value));            //try to parse call value to double
                    else
                        return "!(value)";                                          //parse error was found
                }

                cell.mTree.SetVar(var, value);
            }

            return cell.mTree.Eval().ToString();                                    //evaluate tree after variables have been assigned values
        }

        //description: push undo command onto undo stack and clear the redo stack
        //parameter: command
        //return: 
        public void PushUndo(Cell undoCell, string property)
        {
            switch(property)                                        //use property to determine what ICmd to push
            {
                case "Text":
                    mHistory.PushUndo(new RestoreText(undoCell));
                    break;

                case "BGColor":
                    mHistory.PushUndo(new RestoreColor(undoCell));
                    break;
            }
        }

        //description: push undo command onto undo stack and clear the redo stack
        //parameter: command
        //return: 
        public void PushUndo(List<Cell> undoCells, string property)
        {
            switch (property)                                               //use property to determine what ICmd to push
            {
                case "BGColor":
                    mHistory.PushUndo(new MultiCmds(undoCells, property));
                    break;
            }
        }

        #region History Help Functions
        //description: execute command on top of undo stack and push the inverse command onto the redo stack
        //parameter: 
        //return: 
        public void Undo()
        {
            mHistory.Undo();
        }

        //description: execute command on top of redo stack and push the inverse command onto the undo stack
        //parameter: 
        //return: 
        public void Redo()
        {
            mHistory.Redo();
        }

        //description: checks type of undo command on top of stack
        //parameter: 
        //return: property string
        public string PeekUndo()
        {
            return mHistory.PeekUndo();
        }

        //description: checks type of redo command on top of stack
        //parameter: 
        //return: property string
        public string PeekRedo()
        {
            return mHistory.PeekRedo();
        }

        //description: return number of commands within undo stack
        //parameter: 
        //return: number of commands within undo stack
        public int SizeUndo()
        {
            return mHistory.SizeUndo();
        }

        //description: return number of commands within redo stack
        //parameter: 
        //return: number of commands within redo stack
        public int SizeRedo()
        {
            return mHistory.SizeRedo();
        }
        #endregion

        #region File Loaf/Save Helper Functions
        //description: check if cell is default
        //parameter: cell
        //return: bool
        public bool IsDefaultCell(Cell cell)
        {
            if (cell.Text != "" || cell.BGColor != 4294967295)
                return false;

            else
                return true;
        }

        //description: default all cells
        //parameter: 
        //return: 
        public void DeleteSpreadsheet()
        {
            foreach(Cell delCell in mCells)     //default all cells
            {
                delCell.Text = "";
                delCell.BGColor = 4294967295;
            }
        }
        #endregion

        //description: load stream as xdocument and then into the spreadsheet
        //parameter: stream
        //return: 
        public void LoadFromFile(Stream loadFile)
        {
            XDocument mFile = XDocument.Load(loadFile);                                     //load stream into xdocument

            var elemCells = from Cell in mFile.Root.Element("Cells").Descendants("Cell")    //create list of anonymous type of cells from xdocument
                            select new
                            {
                                Col = int.Parse(Cell.Attribute("Col").Value),               //create fields for the anonymous type with the cell properties
                                Row = int.Parse(Cell.Attribute("Row").Value),
                                Text = Cell.Element("Text").Value,
                                BGColor = uint.Parse(Cell.Element("BGColor").Value)

                            };

            DeleteSpreadsheet();

            foreach(var elemCell in elemCells)                          //iterate through all cells from xdocument
            {
                Cell loadCell = GetCell(elemCell.Col, elemCell.Row);    //determine logic cell

                loadCell.Text = elemCell.Text;
                loadCell.BGColor = elemCell.BGColor;
            }

            #region Other implementation: does not delete all cells first / uses FirstorDefault() many times
            //foreach(Cell loadCell in mCells)                                                //for all cells in the spreadsheet
            //{
            //    var elemCell = elemCells.FirstOrDefault(x => 
            //                                            loadCell.ColumnIndex == x.Col && 
            //                                            loadCell.RowIndex == x.Row);        //attempt to access the cell from the xdocument

            //    if (elemCell != null)                                                       //if it exists load the data into the cell
            //    {
            //        loadCell.Text = elemCell.Text;
            //        loadCell.BGColor = elemCell.BGColor;
            //    }

            //    else                                                                        //else default the cell
            //    {
            //        loadCell.Text = "";
            //        loadCell.BGColor = 4294967295;
            //    }
            //}
            #endregion

            mHistory.Reset();
        }

        //description: save spreadsheet into a xdocument and then into a stream
        //parameter: stream
        //return: 
        public void SaveToFile(Stream saveFile)
        {
            XDocument mFile = new XDocument();
            mFile.Add(new XElement("Spreadsheet"));
            mFile.Element("Spreadsheet").Add(new XElement("Cells"));

            foreach (Cell saveCell in mCells)
            {
                if (!IsDefaultCell(saveCell))
                    mFile.Root.Element("Cells").Add(new XElement("Cell", new XAttribute("Col", saveCell.ColumnIndex), new XAttribute("Row", saveCell.RowIndex),
                                                new XElement("Text", saveCell.Text),
                                                new XElement("BGColor", saveCell.BGColor)));
            }

            #region Other implemntation to try and use SQL(?)
            //XDocument mFile = new XDocument(new XElement("Spreadsheet",
            //                                             new XElement("Cells",
            //                                                          from saveCell in mCells[0,0]
            //                                                          select new XElement("Cell", new XAttribute("Col", saveCell.ColumnIndex), new XAttribute("Row", saveCell.RowIndex),
            //                                                                              new XElement("Text", saveCell.Text),
            //                                                                              new XElement("BGColor", saveCell.BGColor)))));
            #endregion

            mFile.Save(saveFile);
        }
    }
}
