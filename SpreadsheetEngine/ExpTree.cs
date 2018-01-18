using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public class ExpTree
    {
        public abstract class Node
        {
            //description: a function to be replaced by derived classes
            //parameter: 
            //return: 
            public abstract double Eval();
        }

        public class ConstNode : Node
        {
            double mValue;

            //description: constant node constructor
            //parameter: value
            //return: constant node containing value
            public ConstNode(double newValue)
            {
                mValue = newValue;
            }

            //description: returns value
            //parameter: 
            //return: value
            public override double Eval() { return mValue; }
        }

        public class VarNode : Node
        {
            string mName;
            Dictionary<string, double> mDict;

            //description: variable node constructor
            //parameter: name and dictionary
            //return: variable node containing value found with name inside the dictionary
            public VarNode(string newName, Dictionary<string, double> newDict)
            {
                mName = newName;
                mDict = newDict;
                mDict[mName] = 0;   //defaults to 0
            }

            //description: returns value
            //parameter: 
            //return: value
            public override double Eval() { return mDict[mName]; }

            //description: returns name
            //parameter: 
            //return: name
            public string Name { get { return mName; } }
        }

        public class OpNode : Node
        {
            char mOperator;
            Node mLeft;
            Node mRight;

            //description: operator node constructor
            //parameter: operator, left node, and right node
            //return: operator node containing operator and two children nodes
            public OpNode(char newOperator, Node newLeft, Node newRight)
            {
                mOperator = newOperator;
                mLeft = newLeft;
                mRight = newRight;
            }

            //description: returns value based on operator and child values
            //parameter: 
            //return: value returned by operator on two children values
            public override double Eval()
            {
                switch (mOperator)
                {
                    case '+': return mLeft.Eval() + mRight.Eval();
                    case '-': return mLeft.Eval() - mRight.Eval();
                    case '*': return mLeft.Eval() * mRight.Eval();
                    case '/': return mLeft.Eval() / mRight.Eval();

                    default: return 0;
                }
            }

            public Node Left { get { return mLeft; } }
            public Node Right { get { return mRight; } }
        }

        string mExp;
        Node mRoot;
        Dictionary<string, double> mDict;


        //description: expression tree constructor
        //parameter: expression
        //return: expression tree 
        public ExpTree(string newExp = "0")
        {
            mExp = newExp.Replace(" ", "");
            mDict = new Dictionary<string, double>();
            mRoot = Compile(mExp, mDict);
        }

        //description: set variable name to a value withint the dictionary
        //parameter: variable name and value
        //return: 
        public void SetVar(string varName, double varValue)
        {
            mDict[varName] = varValue;
        }

        //description: evaluates the expression tree
        //parameter: 
        //return: result of expression
        public double Eval()
        {
            return mRoot.Eval();
        }

        //description: constructs a node based on the expression passed in
        //parameter: expression and dictionary
        //return: new node
        private static Node BuildSimple(string exp, Dictionary<string, double> dict)
        {
            double num = 0.0;

            if (double.TryParse(exp, out num))
            {
                return new ConstNode(num);      //if exp is a number then return constant with value
            }

            else
                return new VarNode(exp, dict);  //otherwise convert exp to a value via the dictionary 
        }

        //description: determine index of lowest priority operator within expression
        //parameter: expression
        //return: index of lowest priority operator
        private static int GetLowOpIndex(string exp)
        {
            int parenCount = 0;
            int index = -1;                                     //set to -1 to flag if expression is a value or variable
                                                                //   AKA not operator found

            for (int i = exp.Length - 1; i >= 0; i--)            //iterate through expression from right to left
            {
                switch (exp[i])
                {
                    case '(':                                   //increment parenthesis count
                        parenCount++;
                        break;
                    case ')':                                   //decrement parenthesis count
                        parenCount--;
                        break;
                    case '+':                                   //if + or - found and parenthis count is 0
                    case '-':                                   //   then return the current index
                        if (parenCount == 0)
                            return i;
                        break;
                    case '*':                                   //if * or / found and parenthis count is 0
                    case '/':                                   //   and lowest priorty operator not found
                        if (parenCount == 0 && index == -1)     //   then set index to current index
                            index = i;
                        break;
                }
            }
            //return lowest priority operator
            return index;                                       //NOTE: used when lowest priority operator is not a + or -
        }

        //description: construct expression tree based on expression and dictionary
        //parameter: expression and dictionary
        //return: node
        private static Node Compile(string exp, Dictionary<string, double> dict)
        {
            exp = exp.Replace(" ", "");                                 //remove all spaces within expression

            if (exp[0] == '(')                                          //determine if expression is encapsulated by parenthesis
            {                                                           //   if so recompile with their removal
                int parenCount = 1;

                for (int i = 1; i < exp.Length; i++)                    //iterate through expression
                {
                    switch (exp[i])
                    {
                        case '(':                                       //increment parenthesis count
                            parenCount++;
                            break;
                        case ')':                                       //decrement parenthesis count
                            parenCount--;
                            if (parenCount == 0)
                            {
                                if (i == exp.Length - 1)                //if parenthesis count reaches 0 at the end, recompile with the removal of global parenthesis
                                    return Compile(exp.Substring(1, exp.Length - 2), dict);

                                else                                    //if parenthesis count reaches 0 before the end, break
                                    i = exp.Length;                     //   NOTE: differentiates (Hello) + (World) AND (Hello + World)
                            }
                            break;
                    }
                }
            }

            int index = GetLowOpIndex(exp);                             //determine lowest priority operator index

            if (index != -1)                                            //determine if expression contains an operator
            {
                return new OpNode                                       //construct operator node and continue compilation
                    (exp[index],
                    Compile(exp.Substring(0, index), dict),
                    Compile(exp.Substring(index + 1), dict));
            }

            return BuildSimple(exp, dict);                              //construct constant/variable node and continue compilation
        }

        //description: traverse through the tree to return a hashet of the variable names
        //parameter: root node and hashset
        //return: 
        public void GetVarsHelper(Node n, HashSet<string> vars)
        {
            if (n == null)
                return;

            if(n is VarNode)                //if variable node add variable name to hashset
            {
                VarNode var = n as VarNode;

                vars.Add(var.Name);
            }

            else if (n is OpNode)           //if operator node recurse into left and right node
            {
                OpNode op = n as OpNode;

                GetVarsHelper(op.Left, vars);
                GetVarsHelper(op.Right, vars);
            }
        }

        //description: use GetVars starting from root to collect all variable nodes
        //parameter: root node
        //return: hashset of variable names
        public HashSet<string> GetVars()
        {
            HashSet<string> vars = new HashSet<string>();

            GetVarsHelper(mRoot, vars); //search through expression tree for variable names

            return vars;            //return hashset of all variables found within in tree
        }
    }
}