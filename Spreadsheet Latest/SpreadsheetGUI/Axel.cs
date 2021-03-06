﻿using SpreadsheetUtilities;
using SS;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadsheetGUI
{

    /// <summary>
    /// 
    /// </summary>
    public partial class Axel : Form
    {

        private Spreadsheet model; // Keeps track of all the data
        private string FileLocation; // Used for the save function.. stores last file location
        private string V; // For future reference if we want to change the version
        private ArrayList undo;
        private ArrayList undoSpot; 
        /// <summary>
        /// Our Axel Spreadsheet Needs
        /// </summary>
        public Axel()
        {
            InitializeComponent();
            undo = new ArrayList();
            undoSpot = new ArrayList(); // We need to array lists to keep track of position as well as content
            UndoMenuItem.Enabled = false; // Starts out as false
            this.Text = FileLocation; // Makes this equal to a file location
            saveToolStripMenuItem.Enabled = false; // Cannot save until this is enabled
            this.KeyPreview = true; // Get Keyboard Input
            spreadsheetPanel1.SelectionChanged += displaySelection; // Sends the displaySelection delegate as the funciton to be called whenever the spreadsheet changes
            spreadsheetPanel1.SetSelection(0, 1); // These are the cells that start off being selected
            V = "ps6";
            CellName.Text = "A2";
            ValueText.Text = ""; // These are the default settings on how the spreadsheet starts out
            model = new Spreadsheet(s => true, s => s.ToUpper(), V); // Creates a new spread with the versionm ps6
        }

        /// <summary>
        /// Everytime the display is changed in the Spreadsheet panel thsi method is called
        /// </summary>
        /// <param name="ss"></param>
        private void displaySelection(SpreadsheetPanel ss)
        {
            CellContents.Focus(); // Makes sure to focus text to the cell contents text box so that cells recieve input from there
            model.SetContentsOfCell(CellName.Text, CellContents.Text); // Solidifies the text before changing it
            int row, col;
            String value;
            ss.GetSelection(out col, out row); // Gets the values of the cell we are currently selecting
            ss.GetValue(col, row, out value); // Gets the value of the items that we are currently selecting
            string name = GetCellName(col, row); // Gets the name of the cell that we want to change
            CellName.Text = name; // Sets the label according to the name it is in
            try
            {
                ValueText.Text = model.GetCellValue(name).ToString(); // Change the value to the cell if the cursor moves
                CellContents.Text = model.GetCellContents(name).ToString(); // Switch the Cell Contents if it has been moved
                CellContents.SelectionStart = CellContents.Text.Length;
                CellContents.Select(CellContents.Text.Length, 0);
                ss.SetValue(col, row, model.GetCellValue(name).ToString());
            }
            catch
            {

            }
        }

        /// <summary>
        /// Converts a row and column to a name
        /// </summary>
        /// <param name="col_">Will be a character</param>
        /// <param name="row_"> will end up being a row number</param>
        /// <returns></returns>
        private string GetCellName(int col_, int row_)
        {
            string cellName;
            char column = (Char)(col_ + 65); // Changes to ascii code based on character
            return cellName = column + "" + (row_ + 1);
        }

        /// <summary>
        /// Gets the coordinates based on the name of the cell
        /// </summary>
        /// <param name="name"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        private void getCoordinates(string name, out int col, out int row)
        {
            col = (int)name[0] - 65; // Gets the column number
            String num = name.Substring(1);
            double result;
            Double.TryParse(num, out result);
            row = (int)result - 1; // Gets the row number by index
        }

        /// <summary>
        /// When new is clicked a new GUI form is opend up
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewMenuItem_Click(object sender, EventArgs e)
        {
            // Tell the application context to run the form on the same
            // thread as the other forms.
            DemoApplicationContext.getAppContext().RunForm(new Axel());
        }

        /// <summary>
        /// Whenever the text is changed the cell contents box 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CellContents_TextChanged(object sender, EventArgs e)
        {
            int row, col;
            try // SetContentsOfCell may return errors depending on user input
            {
                getCoordinates(CellName.Text, out col, out row); // Gets the coordinates so that we can change that cell
                spreadsheetPanel1.SetValue(col, row, CellContents.Text); // Sets the value of panel as we type into the contents text box
                //displaySelection(spreadsheetPanel1);
                //if (!(model.GetCellValue(GetCellName(col, row)) is FormulaError)) { }
                //ValueText.Text = model.GetCellValue(GetCellName(col, row)).ToString(); // Changes the value text b ox to display the correct value
                undo.Add(CellContents.Text);
                undoSpot.Add(CellName.Text);
                UndoMenuItem.Enabled = true;
            }
            catch
            {

            }
        }

        /// <summary>
        /// This method is called when a key is pressed.. if the key is a enter key then we perform the
        /// compute operations
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CellContents_KeyDown(object sender, KeyEventArgs e)
        {
            int row, col;
            spreadsheetPanel1.GetSelection(out col, out row); // Finds out the current position
  

            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Right || e.KeyCode == Keys.Left || e.KeyCode == Keys.Down || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
            {
                try
                {
                    model.SetContentsOfCell(CellName.Text, CellContents.Text);
                }
                catch
                {
                    CellContents.Text = "%Invalid";
                    MessageBox.Show("A 'Circular Dependency' or 'Invalid Operation' has been detected!", "ERROR",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                try
                {
                    int row_, col_;
                    getCoordinates(CellName.Text, out col_, out row_); // Gets the coordinates so that we can change that cell
                    HashSet<string> names = (HashSet<string>)model.SetContentsOfCell(CellName.Text, CellContents.Text);
                    recalculateCells(names);
                    spreadsheetPanel1.SetValue(col_, row_, model.GetCellValue(CellName.Text).ToString());
                    spreadsheetPanel1.SetSelection(col_, row_ + 1);
                    displaySelection(spreadsheetPanel1); // Change the display on the spreadsheet panel
                    if (model.GetCellContents(CellName.Text) is FormulaError)
                    {
                        CellContents.Text = "%Invalid";
                        MessageBox.Show("A Formular Error Has Occurred!", "ERROR",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch
                {
                    CellContents.Text = "%Invalid";
                    MessageBox.Show("A 'Circular Dependency' or 'Invalid Operation' has been detected!", "ERROR",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                displaySelection(spreadsheetPanel1);

                switch (e.KeyCode)
                {
                    case Keys.Up:
                        spreadsheetPanel1.SetSelection(col, row - 1); // If up key is pressed move up one space
                        break;
                    case Keys.Down:
                        spreadsheetPanel1.SetSelection(col, row + 1); // If down move one down 
                        break;
                    case Keys.Right:
                        spreadsheetPanel1.SetSelection(col + 1, row);
                        break;
                    case Keys.Left:
                        spreadsheetPanel1.SetSelection(col - 1, row);
                        break;
                    default:
                        break;
                }

                e.SuppressKeyPress = true;

                displaySelection(spreadsheetPanel1); // Change the display on the spreadsheet panel
        
             
            }

            if (e.KeyCode == Keys.Delete)
            { // If backspace is selected
                int col_, row_;
                spreadsheetPanel1.GetSelection(out col_, out row_);
                model.SetContentsOfCell(GetCellName(col_, row_), "");
                CellContents.Text = ""; // Makes it null
                displaySelection(spreadsheetPanel1);
            }

        }

        /// <summary>
        /// With this method we will recalucate
        /// </summary>
        /// <param name="cellsToRecalculate"></param>
        private void recalculateCells(ISet<string> cellsToRecalculate)
        {
            foreach (string name in cellsToRecalculate)
            { // Loops through all the cells to recalculate
                Object change = model.GetCellContents(name);
                int col, row;
                model.SetContentsOfCell(name, change.ToString()); ; // Recalculates each cell passed in that is needed to be recalculated
                getCoordinates(name, out col, out row);
                spreadsheetPanel1.SetValue(col, row, model.GetCellValue(name).ToString());
            }
        }
        /// <summary>
        /// This is listening for when the save as button is clicked... Once clicked the spreadsheet should
        /// direct it towards a savedialogue box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveAsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog.Filter = "Spreadsheet Files (.sprd)|*.sprd|All Files (*.*)|*.*";
                saveFileDialog.ShowDialog(); // Opens save dialog
            }
            catch
            {
                saveToolStripMenuItem.Enabled = false;

            }
        }

        /// <summary>
        /// Allows the user to choose a location froma dialogue and save the spreadsheet to that file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            saveToolStripMenuItem.Enabled = true;
            saveMethod(); // Applies the save method
        }

        /// <summary>
        /// This method savs yoru data!
        /// </summary>
        private void saveMethod()

        {
            FileLocation = saveFileDialog.FileName; // Aquires the name that the user wants to save to
            try // Tries to save the document
            {
                model.Save(FileLocation); // Writes as an XML to the file location
            }
            catch // If saved fails then display error dialoge
            {
                saveToolStripMenuItem.Enabled = false; // If save successful takes away the save button until a change is made
                MessageBox.Show("Your File Could Not Be Saved", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// This is the open file menu when clicked on will open a dialog for the user to choose the file to be opened into the spreadsheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Spreadsheet Files (.sprd)|*.sprd|All Files (*.*)|*.*"; // Filter out on xml documents
            openFileDialog.ShowDialog(); // Shows dialogue to the user
        }
        /// <summary>
        /// Once the user chooses a file then the listener will detect it here and we can open that file to a new spreadsheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            FileLocation = openFileDialog.FileName;
            Spreadsheet opened = new Spreadsheet(FileLocation, s => true, s => s.ToUpper(), "ps6"); // Creates a new spreadsheet according to the file location
            int col_, row_;
            foreach (string name in model.GetNamesOfAllNonemptyCells().ToList())
            { // Get rid of old contents
                getCoordinates(name, out col_, out row_); // Gets the coordinates of the area
                spreadsheetPanel1.SetValue(col_, row_, "");
            }
            spreadsheetPanel1.SetSelection(0, 0); // Resets position
            CellName.Text = ""; // Sets the cell text accordin to new spot
            ValueText.Text = ""; // Gets value accordin to the cell content stored in the current spot
            model = opened; // Directs our model to this new data
            int row, col; // coordinates for setting into the spreadsheet
            foreach (string name in opened.GetNamesOfAllNonemptyCells().ToList()) // Gets every name in the cell and then 
            {
                Object contents = opened.GetCellContents(name); // Gets the contents of each name
                Object value = opened.GetCellValue(name); // Gets the value so we can set it in areas
                getCoordinates(name, out col, out row); // Gets the coordinates of the area
                spreadsheetPanel1.SetValue(col, row, value.ToString());
                spreadsheetPanel1.SetSelection(col, row);
                CellName.Text = name; // Sets the cell text accordin to new spot
                ValueText.Text = value.ToString(); // Gets value accordin to the cell content stored in the current spot
                CellContents.Text = contents.ToString(); // Switch the Cell Contents if it has been moved
            }

        }

        /// <summary>
        /// If the spreadsheet changes and the saved command has been changed then this option opens up
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try // Tries to save the document
            {
                model.Save(FileLocation); // Writes as an XML to the file location
            }
            catch // If saved fails then display error dialoge
            {
                saveToolStripMenuItem.Enabled = false; // If not succesful let it go true 
                MessageBox.Show("Your File Could Not Be Saved", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// This will display a help message once clicked!
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Johnathon, I have a special task for you at hand. This spreadsheet works where you are able to select stuff with your mouse. After you select a cell with your mouse you can type in words with your keyboard. You can enter double values as well as formulas by initilizing the use of the word =. You are aloud to save your document in the file. Also you are able to open files. It is pretty sweet dude. Also some special features are that you can use the arrow buttons to move your cursor. Johnathon this mission is dangerous. If you choose to accept it you may not come back.  The spreadsheet may blow up if you are using it wrong so please be careful. Johnathon you are the chosen one. You have been chosen to save the world from bad spreadsheets. More special features include: THE ABILITY TO SAVE WHEN CLOSING! CHANGE THE FONT OF THE SPREADSHEET - THE ABILITY TO USE A UNDO BUTTON!! THE ABILITY TO USE ARROWS TO MOVE!!! THE ABILITY TO SAVE VS SAVE AS!!!! JOHNATHON, THIS IS HUGE you CAN do IT! Please don't be afraid you have been training yoru whole life for this.", "Greetings Good Sir.ALSO REMEMBER DUDE, THAT an important special feature is the ability to SAVE before quitting!! YESS!!! is that not neaT?",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Closes the spreadsheet when the file button is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close(); // Simply Closes the spreadsheet if this is clicked
        }

        /// <summary>
        /// This is what happens when the user attempts to close  in any way
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Axel_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (model.Changed == true) // If the model has been changed prompt user to not close
            {
                DialogResult option = MessageBox.Show("Are you sure you want to close? You will lose unsaved changes! Are you sure you want to quit? Do you want to save your data before quitting??", "STOOOOP!!!", MessageBoxButtons.YesNoCancel);
                if (option.Equals(DialogResult.Cancel))
                {
                    e.Cancel = true; // cancels result
                }
                else if (option.Equals(DialogResult.Yes)) // If it selects yes then saves the spreadsheet
                {
                    saveFileDialog.Filter = "Spreadsheet Files (.sprd)|*.sprd|All Files (*.*)|*.*";
                    saveFileDialog.ShowDialog();
                    saveMethod();
                }

            }

        }

        /// <summary>
        /// Shows really pretty FONTS!!!!! 
        /// YES!!!
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void fontsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    spreadsheetPanel1.Font = fontDialog.Font;
                }
            }
            catch
            {
                MessageBox.Show("HELLO! only 'TRUE Type' Fonts are supported right now!", "SORRY", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Invariant is that both arraylists have count greater then 0
        /// 
        /// This function is called whenever undo button is click and will revert to previous undo spot
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UndoMenuItem_Click(object sender, EventArgs e)
        {
            if(undo.Count == 1){ // If there is only one item left in undo button remove it
                UndoMenuItem.Enabled = false;
            }
            String value = undo[(undo.Count - 1)].ToString();
            String name = undoSpot[(undoSpot.Count - 1)].ToString();
            CellName.Text = name; // We want to be modifying this name
            int col, row;
            getCoordinates(name, out col, out row);
            try
            {
                spreadsheetPanel1.SetSelection(col, row); // Moves selection back to the are of change
                spreadsheetPanel1.SetValue(col, row, value);
                CellContents.Text = value; //Reverts back to old text
                CellName.Text = name;
                model.SetContentsOfCell(name, value);
            }
            catch
            {
                model.SetContentsOfCell(name, "");
            }
            undo.RemoveAt((undo.Count - 1)); // Removes the occurance out of the GUI
            undoSpot.RemoveAt((undoSpot.Count - 1));
        }
    }
    }

