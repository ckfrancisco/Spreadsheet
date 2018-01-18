using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public interface ICmd
    {
        ICmd Exec();
        string Property { get; }
    }

    class History
    {
        private Stack<ICmd> StackUndo;
        private Stack<ICmd> StackRedo;

        //description: constructor to initialize stacks
        //parameter: 
        //return: 
        public History()
        {
            StackUndo = new Stack<ICmd>();
            StackRedo = new Stack<ICmd>();
        }

        //description: clear both stacks of icmds
        //parameter: 
        //return: 
        public void Reset()
        {
            StackUndo.Clear();
            StackRedo.Clear();
        }

        //description: push undo command onto undo stack and clear the redo stack
        //parameter: command
        //return: 
        public void PushUndo(ICmd newUndo)
        {
            StackRedo.Clear();
            StackUndo.Push(newUndo);
        }

        //description: execute command on top of undo stack and push the inverse command onto the redo stack
        //parameter: 
        //return: 
        public void Undo()
        {
            ICmd undo = StackUndo.Pop();
            StackRedo.Push(undo.Exec());
        }

        //description: execute command on top of redo stack and push the inverse command onto the undo stack
        //parameter: 
        //return: 
        public void Redo()
        {
            ICmd redo = StackRedo.Pop();
            StackUndo.Push(redo.Exec());
        }

        //description: checks type of undo command on top of stack
        //parameter: 
        //return: property string
        public string PeekUndo()
        {
            return StackUndo.Peek().Property;
        }

        //description: checks type of redo command on top of stack
        //parameter: 
        //return: property string
        public string PeekRedo()
        {
            return StackRedo.Peek().Property;
        }

        //description: return number of commands within undo stack
        //parameter: 
        //return: number of commands within undo stack
        public int SizeUndo()
        {
            return StackUndo.Count;
        }

        //description: return number of commands within redo stack
        //parameter: 
        //return: number of commands within redo stack
        public int SizeRedo()
        {
            return StackRedo.Count;
        }
    }

    public class RestoreText : ICmd
    {
        private Cell mCell;
        private string mText;

        //description: constructor
        //parameter: cell and cell's old text
        //return: 
        public RestoreText(Cell newCell)
        {
            mCell = newCell;
            mText = newCell.Text;
        }

        //description: set to previous text and return inverse command
        //parameter: 
        //return: 
        public ICmd Exec()
        {
            RestoreText invCmd = new RestoreText(mCell);
            mCell.Text = mText;
            return invCmd;
        }

        //description: returns command's property changed string
        //parameter: 
        //return: property string
        string ICmd.Property { get { return "Text"; } }
    }

    public class RestoreColor : ICmd
    {
        private Cell mCell;
        private uint mBGColor;

        //description: constructor
        //parameter: cell and cell's old color
        //return: 
        public RestoreColor(Cell newCell)
        {
            mCell = newCell;
            mBGColor = newCell.BGColor;
        }

        //description: set to previous color and return inverse command
        //parameter: 
        //return: 
        public ICmd Exec()
        {
            RestoreColor invCmd = new RestoreColor(mCell);
            mCell.BGColor = mBGColor;
            return invCmd;
        }

        //description: returns command's property changed string
        //parameter: 
        //return: property string
        string ICmd.Property { get { return "BGColor"; } }
    }

    public class MultiCmds : ICmd
    {
        private List<ICmd> mCmds;

        //description: constructor
        //parameter: list of cells and the property of the cell changed
        //return: 
        public MultiCmds(List<Cell> newCells, string property)
        {
            mCmds = new List<ICmd>();

            switch(property)                                //use property to determine what ICmd to push
            {
                case "BGColor":
                    foreach (Cell cell in newCells)
                        mCmds.Add(new RestoreColor(cell));
                    break;
            }
        }

        //description: execute commands and return the inverse commands in reverse order
        //parameter: 
        //return: 
        public ICmd Exec()
        {
            MultiCmds invCmds = new MultiCmds(null, null);

            mCmds.Reverse();                    //reverse command order

            foreach(ICmd cmd in mCmds)
            { 
                invCmds.mCmds.Add(cmd.Exec());  //execute commands in reverse order and add to inverse commands
            }

            return invCmds;                     //return inverse commands
        }

        //description: returns command's property changed string
        //parameter: 
        //return: property string
        string ICmd.Property { get { return mCmds.First().Property; } }
    }
}
