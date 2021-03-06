﻿// <author>The Program is finished by AARON KC HSU - 00784935</author>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;


namespace SS
{


    /// <summary>
    /// This class represents adding cells to a group of cells where they will map out a spreadsheet
    /// We keep track of only the active cells to represent infinite cells
    /// </summary>
    public class Spreadsheet : AbstractSpreadsheet
    {
        private DependencyGraph CellLinks; // Keeps track of all the dependent connections between cells
        private Dictionary<string, SpCell> ActiveCells; // Keeps track of active cells and their key names
        // The reason we use a dictionary is so that we can get cells out at O(1) time
        private const String CellPattern = @"^[a-zA-Z]+[a-zA-Z0-9]*$";
      

        /// <summary>
        /// This constructor constructs a spreadsheet imposing some isvalid restrictions and normalize restrictions
        /// This will most likely be used when creating a new document in a spreadsheet program
        /// </summary>
        /// <param name="isValid">a validity delegate</param>
        /// <param name="normalize">a normalization delegate </param>
        /// <param name="version">version that the program was saved under</param>
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version) :
            base(isValid, normalize, version)
        {
            Changed = false; // A new opening starts out as false
            CellLinks = new DependencyGraph(); // Contains all the links in the dependency graph
            ActiveCells = new Dictionary<string, SpCell>(); // The Key we store will be the "Address of the cell on the spreadsheet eventually
        }

        /// <summary>
        /// This constructor constructs a spreadsheet imposing some isvalid restrictions and normalize restrictions
        /// 
        /// This will mostly likely be used to create a spreadsheet from a previously saved spreadsheet file
        /// 
        /// 
        /// </summary>
        /// <param name="file">A string representing a file path</param>
        /// <param name="isValid">a validity delegate</param>
        /// <param name="normalize">a normalization delegate </param>
        /// <param name="version">version that the program is made under</param>
        public Spreadsheet(string file, Func<string, bool> isValid, Func<string, string> normalize, string version) :
            this(isValid, normalize, version)
        {
            if (!this.Version.Equals(GetSavedVersion(file)))
            {
                throw new SpreadsheetReadWriteException("The version this document was made under is not supported in this program");
            }
            LoadFile(file, false); // Loads a file and false indicating that we will perform load and not just version check
        }

        /// <summary>
        /// Spreadsheet object that takes no paremeters and is used to represent a spreadsheet with infinite cells
        /// The sheet will not be standardized to have Excel looking graphs
        /// </summary>
        public Spreadsheet()
            : this(s => true, s => s.ToUpper(), "ps6")
        {
        }

        /// <summary>
        /// Searches through all the cells in the spreadsheet that contains 
        /// </summary>
        /// <returns> A IEnumerable type of all string names of cells in the spread sheet</returns>
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            return ActiveCells.Keys; // Returns all the keys stored in the active cells
        }

        /// <summary>
        /// Takes in the name of a cell based on the validity of the name it should return the contents of the cell
        /// 
        /// If name is null or invalid, throws an InvalidNameException.
        /// 
        /// If name is not initilizaed empty string is returned
        /// </summary>
        /// <param name="name">The name of the cell</param>
        /// <returns>It returns a object of all the cell contents </returns>
        public override object GetCellContents(string name)
        {
            name = Normalize(name);
            if(!IsValid(name)){
                throw new InvalidNameException();
            }
            if (name == null || !ActiveCells.ContainsKey(name))
            { // If the key does not exist that means the cell has nothing in it
                if (Regex.IsMatch(name, CellPattern))
                    return ""; // If it is valid on the spreadsheet return empty string
                else
                    throw new InvalidNameException();
            }
            else
            {
                if (ActiveCells[name].content is Formula)
                    return "=" + ActiveCells[name].content; // Returns the content stored in the cell at the name + the equals
                else
                    return ActiveCells[name].content; // Returns the content stored in the cell at the name
            }
        }

        /// <summary>
        /// An InvalidNameException will be thrown in name or invalid are null or invalid
        /// 
        /// Otherwise, the contents of the named cell becomes number.  The method returns a
        /// set consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// </summary>
        /// <param name="name">The name is the address of the cell on the spreadsheet</param>
        /// <param name="number">Is a number that represents a double in the cell</param>
        /// <returns>A set of cells dependent on the cell that is being changed</returns>
        protected override ISet<string> SetCellContents(string name, double number)
        {
            if (name == null || !Regex.IsMatch(name, @"^[a-zA-Z]+[a-zA-Z0-9]*$"))
            { // If the key does not exist that means the cell has nothing in it
                throw new InvalidNameException();
            }
            if (!ActiveCells.ContainsKey(name)) // If the name does not exist then add it to the dictionary
            {
                ActiveCells.Add(name, new SpCell(name, number)); // New Cell is made to store the string
            }
            else // If it is already active just change it instead
            {
                ActiveCells[name] = new SpCell(name, number);
            }
            CellLinks.ReplaceDependees(name, new HashSet<string>()); // Replace current any variables that was stored previously in the links
            // If error is not thrown procees to adding the new cells into the spreadsheet
            return new HashSet<string>(GetCellsToRecalculate(name)); // Returns all the cells that are dependent on the name
        }
        

        /// <summary>
        /// An ArgumentNullException will be thrown if the contents of the string are null
        /// 
        /// An InvalidNameException will be thrown if the name is Invalid or if it is null
        /// 
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, the contents of the named cell becomes text.  The method returns a
        /// set consisting of name plus the names of all other cells whose value depends, 
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        protected override ISet<string> SetCellContents(string name, string text)
        {
            if (text == null) // Text thrown in method can not be null!
            {
                throw new ArgumentNullException();
            }
            if (name == null || !Regex.IsMatch(name, @"^[a-zA-Z]+[a-zA-Z0-9]*$"))
            { // If the key does not exist that means the cell has nothing in it
                throw new InvalidNameException();
            }
            if (!ActiveCells.ContainsKey(name)) // If the name does not exist then add it to the dictionary
            {
                ActiveCells.Add(name, new SpCell(name, text)); // New Cell is made to store the string
            }
            else // If it is already active just change it instead
            {
                ActiveCells[name] = new SpCell(name, text);
            }
            if (text.Equals(""))
            { // If empty string is entered into the cell remove it
                ActiveCells.Remove(name);
            }
            CellLinks.ReplaceDependees(name, new HashSet<string>()); // Replace current any variables that was stored previously in the links
            // If error is not thrown procees to adding the new cells into the spreadsheet
            return new HashSet<string>(GetCellsToRecalculate(name)); // Returns all the cells that are dependent on the name 
        }

        /// <summary>
        /// If the formula parameter is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if changing the contents of the named cell to be the formula would cause a 
        /// circular dependency, throws a CircularException.  (No change is made to the spreadsheet.)
        /// 
        /// Otherwise, the contents of the named cell becomes formula.  The method returns a
        /// Set consisting of name plus the names of all other cells whose value depends,
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// 
        /// Creates a new cell and adds it to the active cells on the spreadsheet
        /// </summary>
        /// <param name="name">This paremeter is the cell location which is the name</param>
        /// <param name="formula">This object represents a formula if it is inserted in the cell</param>
        /// <returns></returns>
        /// <summary>
        /// If the formula parameter is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name is null or invalid, throws an InvalidNameException.
        /// 
        /// Otherwise, if changing the contents of the named cell to be the formula would cause a 
        /// circular dependency, throws a CircularException.  (No change is made to the spreadsheet.)
        /// 
        /// Otherwise, the contents of the named cell becomes formula.  The method returns a
        /// Set consisting of name plus the names of all other cells whose value depends,
        /// directly or indirectly, on the named cell.
        /// 
        /// For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        /// set {A1, B1, C1} is returned.
        /// 
        /// Creates a new cell and adds it to the active cells on the spreadsheet
        /// </summary>
        /// <param name="name">This paremeter is the cell location which is the name</param>
        /// <param name="formula">This object represents a formula if it is inserted in the cell</param>
        /// <returns></returns>
        protected override ISet<string> SetCellContents(string name, Formula formula)
        {
            // Before adding store the old cell
            Object store = GetCellContents(name);
            // Store dependees before replacing them
            IEnumerable<string> storeD = CellLinks.GetDependees(name);

            try // Resets if a circular exception were to be caughtt
            {

                // We want to replace all the previous variables that were stored in cell if there is none then the new variables will be added
                CellLinks.ReplaceDependees(name, formula.GetVariables()); // The Variables in the equation become the dependees of named cell

                // If error is not thrown procees to adding the new cells into the spreadsheet
                if (!ActiveCells.ContainsKey(name)) // If the name does not exist then add it to the dictionary
                {

                        ActiveCells.Add(name, new SpCell(name, formula, formula.Evaluate(LookupCell)));

                }
                else // If it is already active just change it instead
                {
                        ActiveCells[name] = new SpCell(name, formula, formula.Evaluate(LookupCell));
                }
                return new HashSet<string>(GetCellsToRecalculate(name)); // Returns all the cells that are dependent on the name // Returns all the cells that are dependent on the name
            }
            catch
            {
                SetContentsOfCell(name, store.ToString());
                CellLinks.ReplaceDependees(name, storeD); // Reset everything
                throw new CircularException(); // Throw it!
            }
        }

        /// <summary>
        /// If name is null, throws an ArgumentNullException.
        /// 
        /// Otherwise, if name isn't a valid cell name, throws an InvalidNameException.
        /// 
        /// Otherwise, returns an enumeration, without duplicates, of the names of all cells whose
        /// values depend directly on the value of the named cell.  In other words, returns
        /// an enumeration, without duplicates, of the names of all cells that contain
        /// formulas containing name.
        /// 
        /// For example, suppose that
        /// A1 contains 3
        /// B1 contains the formula A1 * A1
        /// C1 contains the formula B1 + A1
        /// D1 contains the formula B1 - C1
        /// The direct dependents of A1 are B1 and C1
        /// 
        /// </summary>
        /// <param name="name">The name of the cell of whose dependents we want</param>
        /// <returns>The set of all </returns>
        protected override IEnumerable<string> GetDirectDependents(string name)
        {
            return CellLinks.GetDependents(name); // Return the list of direct dependents based on what is in the dependency graph
        }

        /// <summary>
        /// A LookUp Method used to look up a "variable" cell and then return a value for it
        /// 
        /// This will be used to call the formula evaluator when using the constructor of the SpCell class
        /// </summary>
        /// <returns></returns>
        private Object EvaluateCell(string name){
            if (ActiveCells.ContainsKey(name)) // Searches the active cells for a key with the same name
            {
                return ActiveCells[name].Value;
            }
            else // If it doesnt exist the value of the cell should be zero
                return 0.0;
        }

        /// <summary>
        /// Driver method for Evaulate cell used to define the lookup method for the formula evaluator
        /// </summary>
        /// <param name="name">The name of the cell</param>
        /// <returns>a double defined for the variable</returns>
        
        private double LookupCell(string name){

            if (GetCellValue(name) is double)
                return (Double)GetCellValue(name);
                //if (GetCellValue(name).Equals(""))
                    //return 0; // If the cell is nonexistent then return 0
               // else
                   throw new ArgumentException(); // If it is not a string then it is invalid
        }

        /// <summary>
        /// If the spreadsheet has been setted since the opening or saving the spreadsheet then the state of changed will be true
        /// </summary>
        public override bool Changed
        {
            get;
            protected set;
        }

        /// <summary>
        /// Gets the saved version of the spreadsheet
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>Gets the saved version of the spreadsheet</returns>
        public override string GetSavedVersion(string filename)
        {
            return LoadFile(filename, true); // We will get the version using the laod file method
        }

        /// <summary>
        /// 
        /// SpreadsheetReadWriteException will be thrown if your file cannot be saved!
        /// Writes the contents of this spreadsheet to the named file using an XML format.
        /// The XML elements should be structured as follows:
        /// 
        /// <spreadsheet version="version information goes here">
        /// 
        /// <cell>
        /// <name>
        /// cell name goes here
        /// </name>
        /// <contents>
        /// cell contents goes here
        /// </contents>    
        /// </cell>
        /// 
        /// </spreadsheet>
        /// 
        /// There should be one cell element for each non-empty cell in the spreadsheet.  
        /// If the cell contains a string, it should be written as the contents.  
        /// If the cell contains a double d, d.ToString() should be written as the contents.  
        /// If the cell contains a Formula f, f.ToString() with "=" prepended should be written as the contents.
        /// 
        /// If there are any problems opening, writing, or closing the file, the method should throw a
        /// SpreadsheetReadWriteException with an explanatory message.
        /// <summary/>
        /// <param name="filename"></param>
        public override void Save(string filename)
        {
            // Tries to saved
            try { 
                Changed = false; // Assuming save does not fail
            using (XmlWriter writer = XmlWriter.Create(filename)) // Creates a Xml Document with the current filename
            {
                writer.WriteStartDocument(); // YES! This starts the document
                writer.WriteStartElement("spreadsheet"); // 
                writer.WriteAttributeString("version", this.Version); // Adds the version of the spreadsheet  
                foreach (string n in ActiveCells.Keys) // Write cells
                {
                    if (ActiveCells[n].content is Formula) // If it is a formula we need to add the = to differentiate it into being a formula
                    {
                        string formula = "=" + ActiveCells[n].content.ToString(); 
                        writer.WriteStartElement("cell"); // A Cell element is created to contain the name and content
                        writer.WriteElementString("name", n); // Gives the name of the cell
                        writer.WriteElementString("contents", formula);
                        writer.WriteEndElement(); // </cell>
                    }
                    else // If it is a double or string we can just add the content.ToString() to the data file
                    {
                    writer.WriteStartElement("cell"); // A Cell element is created to contain the name and content
                    writer.WriteElementString("name", n); // Gives the name of the cell
                    writer.WriteElementString("contents", ActiveCells[n].content.ToString());
                    writer.WriteEndElement(); // </cell>
                    }
                }
                writer.WriteEndElement(); // </spreadsheet>
                writer.WriteEndDocument(); // Exists writer
            }
            }
                catch{
                    Changed = true; // Since saved failed it is still in saved state
                    throw new SpreadsheetReadWriteException("Your file could not be saved");
                }
           }
        

        /// <summary>
        /// Each cell may represent a value according to the content stored into the cell
        /// This method checks a cell by name and obtains that value
        /// 
        /// If the content is a string it returns a string
        /// If it is double it returns the double
        /// If it is a formula it evaluates the formula
        /// 
        /// </summary>
        /// <param name="name"> This is the name of the cell</param>
        /// <returns></returns>
        public override object GetCellValue(string name)
        {
            name = Normalize(name);
            if (!IsValid(name))
            {
                throw new InvalidNameException();
            }
            if (name == null || !ActiveCells.ContainsKey(name))
            { // If the key does not exist that means the cell has nothing in it
                if (Regex.IsMatch(name, @"^[a-zA-Z]+[a-zA-Z0-9]*$"))
                    return ""; // If it is valid on the spreadsheet return empty string
                else
                    throw new InvalidNameException();
            }
            else
                return ActiveCells[name].Value; // Returns the value of the content stored in the cell at the name
        }

        /// <summary>
        /// This method sets what ever the user inputs into the cells... 
        /// If the user uses the = command it will be a formula
        /// If the user enters a double it recognizes it as a double
        /// If the user enters a string it also recognizes it
        /// </summary>
        /// <param name="name">name of the cell like A1</param>
        /// <param name="content">the stuff written in the cell</param>
        /// <returns>A set of all the cells that needs to be changed</returns>
        public override ISet<string> SetContentsOfCell(string name, string content)
        {
         name = Normalize(name);
         if (!IsValid(name)) { // Name must be valid in spreadsheet
             throw new InvalidNameException();
         }
            Changed = true; // When something has been changed then this is true
            Double num; // To store the double if it happens to be a double
            HashSet<String> returnCells = new HashSet<String>();
            if (content.FirstOrDefault().Equals('=')) // If the user types = at the beginning then it is a formula
            {
                content = content.Substring(1); // Removes the =
                returnCells = (HashSet<String>)SetCellContents(name, new Formula(content, this.Normalize, s => this.IsValid(s) && Regex.IsMatch(s, @"[a-zA-Z](?: [a-zA-Z]|\d)*"))); // Set using formula method using the current define is valid and normalize
            }
            else if(Double.TryParse(content, out num)){ // Trys to parse the content
                returnCells = (HashSet<String>)SetCellContents(name, num); // Set using double method
            }
            else // If it is not a double or a formula it is a string
            {
                returnCells = (HashSet<String>)SetCellContents(name, content); // Set using string method
            }
            recalculateCells(returnCells);
            return returnCells;
        }

        /// <summary>
        /// With this method we will recalucate
        /// </summary>
        /// <param name="cellsToRecalculate"></param>
        private void recalculateCells(ISet<string> cellsToRecalculate)
        {
            foreach(string name in cellsToRecalculate){ // Loops through all the cells to recalculate
                Object change = GetCellContents(name);
                if(change is Formula){
                    Formula f = (Formula)change;
                    ActiveCells[name] = new SpCell(name, f, f.Evaluate(LookupCell)); ; // Recalculates each cell passed in that is needed to be recalculated
                }
            }
        }

        /// <summary>
        /// This helper method will be used both in the constructor as well as the get saved version method
        /// 
        /// The boolean indicates the place where this method is used
        /// </summary>
        /// <param name="filename">the name of the file we are trying to get data from</param>
        /// <param name="VersionReturn">Where the method is being called</param>
        /// <returns>a version string</returns>
        private string LoadFile(string filename, Boolean VersionReturn)
        {
            try // Tries to load a file
            {
                Changed = false; // Assuming it works
                using (XmlReader reader = XmlReader.Create(filename)) // Creates a reader from the file name input
                {
                    while (reader.Read())
                    { // As long as there is a attribute to read
                        if (reader.IsStartElement())
                        {
                                switch (reader.Name)
                                {
                                    case "spreadsheet":
                                        if (VersionReturn)
                                            return reader["version"]; // Get the attribute of spreadsheet which is the version
                                        break;
                                    case "cell": // If the element is just a cell do not do anything
                                        break;
                                    case "name":
                                        reader.Read();
                                        string n = reader.Value; // Stores name
                                        reader.Read(); // Goes past skils 
                                        reader.Read(); // Goes to next item which should be contents
                                        reader.Read();
                                        SetContentsOfCell(n, reader.Value); // Set a new cell with the contents of the current reader
                                        break;
                                    default:
                                        throw new SpreadsheetReadWriteException("Invalid Tag");
                                }
                        }
                    }
                }
                return ""; // At this point we do not need to return a version
            }
            catch
            {
                Changed = true; // If load fails then it is still in changed state
                throw new SpreadsheetReadWriteException("Your file could not be loaded");
            }
        }
    }
}
