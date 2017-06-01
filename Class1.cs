/*
Nathan Young
SID# 10932954
CPTS 321
HW 7
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Xml.Linq;
using System.Xml;

namespace SpreadsheetEngine
{
    public class Class1
    {

    }
}


// **************************************************************************************************
// ************************************ SPREADSHEET CELL CLASS **************************************
// **************************************************************************************************


namespace CptS321
{

    public abstract class SpreadsheetCell : INotifyPropertyChanged
    {
        protected int r_index; // row
        protected int c_index; // column
        protected string data; // Text
        protected string thevalue; // Value
        public event PropertyChangedEventHandler PropertyChanged;
        protected uint dacolor; // Background Color
        protected bool isdefault; // Default values check

        public SpreadsheetCell(int column, int row) // constructor
        {
            r_index = row;
            c_index = column;
            data = null;
            dacolor = 4294967295;
            isdefault = true;
        }

        // ****************************** PROPERTIES *************************************

        public int RowIndex
        { get { return r_index; } }

        public int ColumnIndex
        { get { return c_index; } }

        public string Text
        {
            get { return data; }
            set
            {
                data = value;
                PropertyChanged(this, new PropertyChangedEventArgs("Text")); // notify subscriber
            }
        }

        public string Value
        {
            get { return thevalue; }
            internal set { thevalue = value; }
        }

        public uint BGColor
        {
            get { return dacolor; }
            set
            {
                { dacolor = value; }
                PropertyChanged(this, new PropertyChangedEventArgs("BGColor"));
            }
        }

        public bool Default
        {
            get { return isdefault; }
            set { isdefault = value; }
        }
    }




    // **************************************************************************************************
    // *************************************** SPREADSHEET CLASS ***************************************
    // **************************************************************************************************




    public class Spreadsheet
    {
        int rowcount = 0;
        int columncount = 0;
        SpreadsheetCell[,] thelist; // generate array
        public event PropertyChangedEventHandler CellPropertyChanged = delegate { };
        private Dictionary<SpreadsheetCell, HashSet<SpreadsheetCell>> dependancies; // Dependencies Chart
        private Stack<ICmd> Undos;
        private Stack<ICmd> Redos;

        // ********************** BUILD SPREADSHEET *********************

        public Spreadsheet(int columnsize, int rowsize) // constructor
        {
            rowcount = rowsize;
            columncount = columnsize;
            thelist = new StringCell[columnsize, rowsize];

            for (int i = 0; i < columnsize; i++) // generate cells
            {
                for (int j = 0; j < rowsize; j++)
                {
                    SpreadsheetCell newcell = buildcell(CellType.String, i, j); // create cell
                    thelist[i, j] = newcell;                                    // add to array
                    thelist[i, j].PropertyChanged += Spreadsheet_PropertyChanged; // add subscriber
                }
            }
            dependancies = new Dictionary<SpreadsheetCell, HashSet<SpreadsheetCell>>(); // initialize Dependencies
            Undos = new Stack<ICmd>();
            Redos = new Stack<ICmd>();
        }

        //  ********************** SAVE SPREADSHEET *********************

        public void SaveSpreadsheet(Stream thestream)
        {
            XDocument thedoc = new XDocument();

            XElement theroot = new XElement("Spreadsheet");
            foreach (SpreadsheetCell cell in thelist)
            {
                if (cell.Default == false)          // for each cell w/o default values
                { theroot.Add(PackageCell(cell)); } // package cell and add to file
            }
            thedoc.Add(theroot);
            XmlWriter huh = XmlWriter.Create(thestream);
            thedoc.Save(thestream); // send to file

        }

        protected XElement PackageCell(SpreadsheetCell thecell) // Builds package containing cell data
        {
            XElement package = new XElement("cell");
            XElement BGColor = new XElement("bgcolor", thecell.BGColor.ToString("X")); // Store BGColor as Hex
            XElement Text = new XElement("text", thecell.Text);                        // Store Text
            package.Add(BGColor);
            package.Add(Text);
            XName name = "name";
            string cellname = ((char)(thecell.ColumnIndex + 65)).ToString() + (thecell.RowIndex + 1).ToString(); // cell rc-index as name
            package.SetAttributeValue(name, cellname);
            return package;
        }

        // ********************** LOAD SPREADSHEET ***************************

        public void LoadSpreadsheet(Stream thestream)
        {

            SpreadsheetCell thecell = null;
            foreach (SpreadsheetCell cell in thelist) // reset all conditions
            {
                if (cell.Default == false)
                {
                    cell.Text = "";
                    cell.Value = "";
                    cell.BGColor = 4294967295; // <- color white
                    cell.Default = true;
                }
            }
            this.Undos.Clear();
            this.Redos.Clear();

            using (XmlReader huh = XmlReader.Create(thestream))
            {
                while (huh.Read())
                {
                    if (huh.NodeType == XmlNodeType.Element)
                    {
                        if (huh.Name == "cell")
                        {
                            XElement el = XNode.ReadFrom(huh) as XElement; // declare cell element

                            // get cell rc-index
                            string name = el.Attribute("name").Value;
                            thecell = this.GetCell((int)(((char)(name[0])) - 65), Int32.Parse(name.Substring(1)) - 1);

                            // get cell bgcolor
                            string bgcolor = el.Element("bgcolor").Value;
                            bgcolor = bgcolor.Replace(" ", ""); // remove whitespace
                            bgcolor = bgcolor.Replace("\n", "");
                            thecell.BGColor = (uint)Int32.Parse(bgcolor, System.Globalization.NumberStyles.HexNumber); // convert from hex

                            // get cell contents
                            string text = el.Element("text").Value;
                            text = text.Replace("\n", "");
                            if (text != "")
                            { thecell.Text = text; }

                        }
                    }
                }
            }
        }

        // ********************** CELL CHANGE *********************

        private void Spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            SpreadsheetCell thecell = sender as SpreadsheetCell;
            if (e.PropertyName == "Text")
            {
                string celldata = thecell.Text;
                bool circref = false;

                if (celldata != "" && celldata != null && celldata[0] == '=') // check if formula
                {
                    celldata = celldata.Remove(0, 1); // remove '='
                    ExpTree newexpression = new ExpTree(celldata); // build expression tree
                    int column = 0; // initialize values
                    int row = 0;
                    double data = 0;
                    bool isstring = false; // check if referring cell is a string


                    for (int i = 0; i < newexpression.thelist.Count; i++) // for each variable in exptree dictionary
                    {
                        column = (int)newexpression.thelist.ElementAt(i).Key[0] - 65;// get source column
                        if (Int32.TryParse(newexpression.thelist.ElementAt(i).Key.Remove(0, 1), out row) == false)
                        { thecell.Value = "!Invalid Reference"; isstring = true; break; } // Invalid Reference Condition

                        row = (Int32.Parse(newexpression.thelist.ElementAt(i).Key.Remove(0, 1))) - 1;// get source row
                        if (row >= this.rowcount)
                        { thecell.Value = "!Out of Range"; isstring = true; break; } // Out of Range Condition

                        if (column == thecell.ColumnIndex && row == thecell.RowIndex) // Self - Reference condition
                        { thecell.Value = "!Self Reference"; isstring = true; break; }

                        SpreadsheetCell ReferredCell = this.GetCell(column, row); // thecell references this cell

                        if (CheckForCircularReference(thecell, ReferredCell) == true)       // Check for Circular Reference
                        {
                            thecell.Value = "!Circular Reference";
                            isstring = true;
                            circref = true;
                            break;
                        }

                        if (!dependancies.ContainsKey(ReferredCell)) // if first time referred cell has dependency
                        { dependancies.Add(ReferredCell, new HashSet<SpreadsheetCell>()); } // initialize index

                        dependancies[ReferredCell].Add(thecell);                            // add dependency in spreadsheet

                        UpdateDependancies(ReferredCell);                                   // update dependencies

                        if (!Double.TryParse(ReferredCell.Value, out data) && ReferredCell.Value != null) // if content is a string i.e. not a number
                        {
                            thecell.Value = thelist[column, row].Value;
                            isstring = true;
                            break;
                        }
                        else if (ReferredCell.Value == null)                  // if source content is blank
                        { newexpression.SetVar(newexpression.thelist.ElementAt(i).Key, data); }
                        else                                                                // Gather reference values
                        {
                            data = Double.Parse(ReferredCell.Value); // convert content to number
                            newexpression.SetVar(newexpression.thelist.ElementAt(i).Key, data); // assign cell value to variable
                        }
                    }
                    if (isstring == false)
                    { thecell.Value = newexpression.Eval().ToString(); } // evaluate expression

                }
                else
                { thecell.Value = thecell.Text; } // if referred cell is string, copy text
                if (dependancies.ContainsKey(thecell) && circref == false) // update values of all dependent cells
                {
                    foreach (SpreadsheetCell updateit in dependancies[thecell])
                    { updateit.Text = updateit.Text; }
                }
                thecell.Default = false;
                CellPropertyChanged(sender, new PropertyChangedEventArgs("Value")); // notify subsciber
            }

            if (e.PropertyName == "BGColor")
            {
                thecell.Default = false;
                CellPropertyChanged(sender, new PropertyChangedEventArgs("BGColor"));
            }
        }

        void UpdateDependancies (SpreadsheetCell thecell)
        {
            foreach(SpreadsheetCell dep in dependancies[thecell])
            {
                if(dep.Text[0] != '=')
                {
                    dependancies[thecell].Remove(dep);
                }
                else
                {
                    ExpTree depsmama = new ExpTree(dep.Text.Remove(0, 1));
                    string thecellname = ((char)(thecell.ColumnIndex + 65)).ToString() + (thecell.RowIndex + 1).ToString();
                    if (!depsmama.thelist.ContainsKey(thecellname))
                    {
                        dependancies.Remove(dep);
                    }
                }
            }
        }

        bool CheckForCircularReference(SpreadsheetCell thecell, SpreadsheetCell CheckAgainst) // returns true if circ ref exists
        {
            if (!dependancies.ContainsKey(thecell))
                return false;
            if (dependancies[thecell].Contains(CheckAgainst))
                return true;
            else
            {
                foreach (SpreadsheetCell dep in dependancies[thecell])
                {
                    if (CheckForCircularReference(dep, CheckAgainst) == true)
                        return true;
                }
            }
            return false;
        }

        //******************************** UNDO REDO COMMANDS ********************************

        public interface ICmd // base class
        { ICmd Exec(); }

        public class RestoreText : ICmd
        {
            private SpreadsheetCell oldcell;
            private string oldtext;
            public RestoreText(SpreadsheetCell c, string t) // constructor
            { oldcell = c; oldtext = t; }
            public ICmd Exec()
            {
                RestoreText inverse = new RestoreText(oldcell, oldcell.Text);
                oldcell.Text = oldtext;
                return inverse; // return redo
            }
        }

        public class RestoreBGColor : ICmd
        {
            internal SpreadsheetCell oldcell;
            internal uint oldcolor;
            public RestoreBGColor(SpreadsheetCell c, uint b)
            { oldcell = c; oldcolor = b; }
            public ICmd Exec()
            {
                RestoreBGColor inverse = new RestoreBGColor(oldcell, oldcell.BGColor);
                oldcell.BGColor = oldcolor;
                return inverse;
            }
        }

        public class MultiCmdBGColor : ICmd // for multiple cell changes
        {
            RestoreBGColor[] thecommands;
            public ICmd Exec()
            {
                RestoreBGColor[] sub = new RestoreBGColor[thecommands.Length];
                for (int i = 0; i < thecommands.Length; i++)
                {
                    sub[i] = new RestoreBGColor(thecommands[i].oldcell, thecommands[i].oldcell.BGColor);
                    thecommands[i].Exec();
                }
                MultiCmdBGColor inverse = new MultiCmdBGColor(sub);
                return inverse;
            }
            public MultiCmdBGColor(SpreadsheetCell[] c) // constructor for array of cells
            {
                thecommands = new RestoreBGColor[c.Length];
                for (int i = 0; i < c.Length; i++)
                { thecommands[i] = new RestoreBGColor(c[i], c[i].BGColor); }
            }
            public MultiCmdBGColor(RestoreBGColor[] r) // constructor for array of commands
            {
                thecommands = new RestoreBGColor[r.Length];
                for (int i = 0; i < r.Length; i++)
                { thecommands[i] = r[i]; }
            }
        }

        //******************************** PUSH STACK *********************************

        public void AddUndo(SpreadsheetCell c, string t)
        {
            RestoreText thething = new RestoreText(c, t);
            Undos.Push(thething);
        }

        public void AddUndo(SpreadsheetCell c, uint b)
        {
            RestoreBGColor thething = new RestoreBGColor(c, b);
            Undos.Push(thething);
        }

        public void AddUndo(SpreadsheetCell[] c)
        {
            MultiCmdBGColor thething = new MultiCmdBGColor(c);
            Undos.Push(thething);
        }

        // *************************** POP STACK *****************************************

        public string PopUndoStack() 
        {
            ICmd thething = Undos.Pop();
            Redos.Push(thething.Exec()); // push inverse onto redo stack
            if (UndoIsEmpty() == false)
            {
                if (Undos.Peek().GetType().ToString() == "CptS321.Spreadsheet+RestoreText")
                { return "Undo Text Change"; }
            }
            else { return "Undo"; }
            return "Undo Background Color Change";
        }

        public string PopRedoStack() 
        {
            ICmd thething = Redos.Pop();
            Undos.Push(thething.Exec()); // push inverse onto undo stack
            if (RedoIsEmpty() == false)
            {
                if (Redos.Peek().GetType().ToString() == "CptS321.Spreadsheet+RestoreText")
                { return "Redo Text Change"; }
            }
            else { return "Redo"; }
            return "Redo Background Color Change";
        }

        // ********************************** PEEK STACK *********************************

        public string PeekUndoStack()
        {
            if (UndoIsEmpty() == false)
            {
                if (Undos.Peek().GetType().ToString() == "CptS321.Spreadsheet+RestoreText")
                { return "Undo Text Change"; }
            }
            else { return "Undo"; }
            return "Undo Background Color Change";
        }

        public string PeekRedoStack()
        {
            if (RedoIsEmpty() == false)
            {
                if (Redos.Peek().GetType().ToString() == "CptS321.Spreadsheet+RestoreText")
                { return "Redo Text Change"; }
            }
            else { return "Redo"; }
            return "Redo Background Color Change";
        }

        // *********************************** STACK ISEMPTY ******************************

        public bool UndoIsEmpty() // returns true if undo stack is empty
        {
            if (Undos.Count == 0)
            { return true; }
            return false;
        }

        public bool RedoIsEmpty() // returns true if redo stack is empty
        {
            if (Redos.Count == 0)
            { return true; }
            return false;
        }


        public void ClearRedoStack()
        { Redos.Clear(); }

        // ********************************** CELL BUILDER *************************************

        public SpreadsheetCell buildcell(CellType type, int column, int row)  // cell builder
        {
            switch (type)
            {
                case CellType.Formula:
                    return new FormulaCell(column, row);
                case CellType.String:
                    return new StringCell(column, row);
                case CellType.Value:
                    return new ValueCell(column, row);
                default:
                    throw new NotSupportedException();
            }
        }

        public SpreadsheetCell GetCell(int column, int row) // getter functions
        { return thelist[column, row]; }

        public int RowCount
        { get { return rowcount; } }

        public int ColumnCount
        { get { return columncount; } }

    }

    // **************************************************************************************************
    // ******************************* USELESS CLASSES ***********************************************
    // **************************************************************************************************


    public class ValueCell : SpreadsheetCell
    {
        public ValueCell(int column, int row) : base(column, row)
        { }
    }
    public class FormulaCell : SpreadsheetCell
    {
        public FormulaCell(int column, int row) : base(column, row)
        { }
    }
    public class StringCell : SpreadsheetCell
    {
        public StringCell(int column, int row) : base(column, row)
        { }
    }

    public enum CellType
    {
        Value,
        String,
        Formula
    }



    // **************************************************************************************************
    // *********************************** EXPTREE CLASS ********************************************
    // **************************************************************************************************





    public class ExpTree
    {
        public string theexpression;
        public Dictionary<string, double> thelist;
        Node Root = null;

        public ExpTree(string expression)
        {
            theexpression = expression.Replace(" ", ""); // lose white space
            thelist = new Dictionary<string, double>();  // initialize dictionary
            Root = buildtree(theexpression);             // generate tree
        }


        private Node buildtree(string exp) // tree generator
        {
            int i = GetLow(exp); // get lowest priority operator
            if (i > 0) // if found
            {
                string startexp = exp.Remove(i, (exp.Length - i)); // parse front half
                string endexp = exp.Remove(0, i + 1);             // parse tail half
                return new BinaryNode(exp[i],                     // create new node
                    buildtree(startexp),
                    buildtree(endexp));
            }

            else if (exp[0] == '(') // if not found, but expression has parenthesis
            {
                string newexp = exp.Remove(0, 1);
                return buildtree(newexp.Remove(newexp.Length - 1)); // remove parenthesis and try again
            }
            else
            { return buildtreesimple(exp); } // if not an operator, create operand
        }

        private int GetLow(string exp) // get lowest priority operator
        {
            int counter = 0; // parenthesis counter
            int index = -1;

            for (int i = exp.Length - 1; i >= 0; i--) // for each element in string
            {
                switch (exp[i])
                {
                    case ')':
                        counter++;
                        break;
                    case '(':
                        counter--;
                        break;
                    case '+':
                        if (counter == 0)
                            return i;
                        break;
                    case '-':
                        if (counter == 0)
                            return i;
                        break;
                    case '*':
                        if (counter == 0 && index == -1)
                            index = i;
                        break;
                    case '/':
                        if (counter == 0 && index == -1)
                            index = i;
                        break;
                }
            }
            return index;
        }

        private Node buildtreesimple(object data) // determine if constant or variable
        {
            double value;
            if (Double.TryParse((string)data, out value)) // if data is a number
            { return new ConstNode(value); }
            else
            {
                thelist.Add((string)data, 0); // add default '0' value to dictionary
                return new VarNode((string)data);
            }
        }

        public void SetVar(string varName, double varValue) // set variable
        {
            if (!thelist.ContainsKey(varName)) // if name is not in dictionary
            { thelist.Add(varName, varValue); }
            else
            { thelist[varName] = varValue; }
        }

        public double Eval()
        { return Root.Eval(thelist); }

        public abstract class Node // Base class
        {
            public object data;
            public virtual double Eval(Dictionary<string, double> list)
            { return (double)data; }
        }

        public class BinaryNode : Node // operator node
        {
            public Node lchild;
            public Node rchild;
            public BinaryNode(char op, Node left, Node right)
            {
                data = op;
                lchild = left;
                rchild = right;
            }
            public override double Eval(Dictionary<string, double> list)
            {
                switch ((char)data)
                {
                    case '+':
                        return (lchild.Eval(list) + rchild.Eval(list));
                    case '-':
                        return (lchild.Eval(list) - rchild.Eval(list));
                    case '*':
                        return (lchild.Eval(list) * rchild.Eval(list));
                    case '/':
                        return (lchild.Eval(list) / rchild.Eval(list));
                }
                return 0;
            }
        }

        public class VarNode : Node // variable node
        {
            public VarNode(string newvar)
            { data = newvar; }

            public override double Eval(Dictionary<string, double> thelist)
            {
                if (thelist.ContainsKey((string)data)) // look up variable value
                    return thelist[(string)data];
                else
                    Console.WriteLine("ERROR: Unknown Variable Detected: '{0}'", (string)data);
                return 0;
            }
        }

        public class ConstNode : Node // uses base class eval()
        {
            public ConstNode(double newdat)
            { data = newdat; }
        }
    }
}
