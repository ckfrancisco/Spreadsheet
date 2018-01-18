using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using SpreadsheetEngine;

namespace SpreadsheetEngine
{
    public abstract class Cell : INotifyPropertyChanged
    {
        private int mColumnIndex;
        private int mRowIndex;

        protected string mText;
        protected string mValue;
        protected uint mBGColor;

        public ExpTree mTree;

        public event PropertyChangedEventHandler PropertyChanged;

        //description: signals property changed
        //parameters: property changed arguements
        //return: 
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));  //pass cell as object sender
            }
        }

        //description: cell constructor
        //parameters: column index, row index, text, and value
        //return: 
        public Cell(int newColumnIndex = 0, int newRowIndex = 0, string newText = "", string newValue = "")
        {
            mColumnIndex = newColumnIndex;
            mRowIndex = newRowIndex;
            mText = newText;
            mValue = newValue;
            mBGColor = 4294967295;
            mTree = null;
        }

        //description: column property getter
        //parameters: 
        //return: column index
        public int ColumnIndex { get { return mColumnIndex; } }

        //description: row property getter
        //parameters: 
        //return: row index
        public int RowIndex { get { return mRowIndex; } }

        //description: text property getter and setter, setter signals property change
        //parameters: setter requires being set to another string
        //return: getter returns text
        public string Text
        {
            get { return mText; }
            set
            {
                if (mText != value)             //if attempted value change is the same then skip
                {
                    mText = value;
                    OnPropertyChanged("Text");  //signal cell property changed with string indicating the which cell's text was changed
                }
            }
        }

        //description: background color uint getter and setter
        //parameters: setter requires being set to another uint
        //return: getter returns uint
        public uint BGColor
        {
            get { return mBGColor; }
            set
            {
                if (mBGColor != value)              //if attempted value change is the same then skip
                {
                    mBGColor = value;
                    OnPropertyChanged("BGColor");   //signal cell property changed with string indicating the which cell's background color was changed
                }
            }
        }

        //description: value property getter
        //parameters: 
        //return: value
        public string Value { get { return mValue; } }
    }
}