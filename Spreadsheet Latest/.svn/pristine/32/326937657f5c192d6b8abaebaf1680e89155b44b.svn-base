// <author>The Program is finished by AARON KC HSU - 00784935</author>

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SS;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using SpreadsheetUtilities;

namespace SpreadsheetTests
{
    /// <summary>
    /// This class is used to test the baloneys out of the spreadsheet solution and all its containing classes
    /// </summary>
    [TestClass]
    public class SpreadTests
    {
        /// <summary>
        /// A test case that checks valid name function
        /// </summary>
        [TestMethod]
        public void checkName()
        {
            String s = "Ln";
            if (Regex.IsMatch(s, @"[a-zA-Z_](?: [a-zA-Z_]|\d)*"))
            {
                Assert.IsTrue(true);
            }
            else
            {
                Assert.IsTrue(false);
            }
        }

        /// <summary>
        /// A test case that checks valid name function
        /// </summary>
        [TestMethod]
        public void checkName1()
        {
            String s = "___Ln";
            if (Regex.IsMatch(s, @"[a-zA-Z_](?: [a-zA-Z_]|\d)*"))
            {
                Assert.IsTrue(true);
            }
            else
            {
                Assert.IsTrue(false);
            }
        }


        /// <summary>
        /// A test case that checks valid name function
        /// </summary>
        [TestMethod]
        public void checkName3()
        {
            String s = "A8";
            if (Regex.IsMatch(s, @"[a-zA-Z_](?: [a-zA-Z_]|\d)*"))
            {
                Assert.IsTrue(true);
            }
            else
            {
                Assert.IsTrue(false);
            }
        }

        /// <summary>
        /// A test case that checks valid name function
        /// </summary>
        [TestMethod]
        public void checkName4()
        {
            String s = "6A";
            if (Regex.IsMatch(s, @"^[a-zA-Z_](?: [a-zA-Z_]|\d)*$"))
            {
                Assert.IsTrue(false);
            }
            else
            {
                Assert.IsTrue(true);
            }
        }

        /// <summary>
        /// A test case that checks valid name function
        /// </summary>
        [TestMethod]
        public void checkName5()
        {
            String s = "A6";
            if (Regex.IsMatch(s, @"^[a-zA-Z_](?: [a-zA-Z_]|\d)*$"))
            {
                Assert.IsTrue(true);
            }
            else
            {
                Assert.IsTrue(false);
            }
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        public void GetCellContents()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("A1", "5.8");
            contents.GetCellContents("A1");
        }


        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellContents1()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.GetCellContents("+ HHA1");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellContents2()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.GetCellContents("5 + !@#$");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        public void GetCellContents3()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            Assert.AreEqual(contents.GetCellContents("A1"), "");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        public void SetCellContents()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("A1", "5.0");
            Assert.IsTrue(spreadsheet.GetCellContents("A1").Equals(5.0));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests that A1 and a1 should both fit in the cell
        /// </summary>
        [TestMethod]
        public void SetCellContents2()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("A1", "5.0");
            spreadsheet.SetContentsOfCell("A2", "7.9");
            spreadsheet.SetContentsOfCell("a1", "10.0");
            Assert.IsTrue(spreadsheet.GetCellContents("a1").Equals(10.0));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests that A1 can be overidden
        /// </summary>
        [TestMethod]
        public void SetCellContents3()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("A1", "5.0");
            spreadsheet.SetContentsOfCell("A1", "10.0");
            Assert.IsTrue(spreadsheet.GetCellContents("A1").Equals(10.0));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests that strings can be inserter
        /// </summary>
        [TestMethod]
        public void SetCellContents4()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("A1", "text");
            spreadsheet.SetContentsOfCell("A12", "10.0");
            Assert.IsTrue(spreadsheet.GetCellContents("A1").Equals("text"));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests that single characters are valid variables
        /// </summary>
        [TestMethod]
        public void SetCellContents5()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("x", "text");
            spreadsheet.SetContentsOfCell("A12", "10.0");
            Assert.IsFalse(spreadsheet.GetCellContents("A12").Equals("text"));
            Assert.IsTrue(spreadsheet.GetCellContents("x").Equals("text"));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests that single characters are valid variables
        /// </summary>
        [TestMethod]
        public void SetCellContents6()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("x", "text");
            spreadsheet.SetContentsOfCell("A12", "10.0");
            Assert.IsTrue(spreadsheet.GetCellContents("x").ToString().Equals("text"));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This teststhe first case of
        /// </summary>
        [TestMethod]
        public void SetCellContents7()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("x", "text");
            IEnumerable<string> set = spreadsheet.SetContentsOfCell("A12", "10.0");
            foreach(string s in set)
            Assert.IsTrue(s.Equals("A12"));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// </summary>
        [TestMethod]
        public void SetCellContentsSet()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "5.0");
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("C1+A1").ToString()); // A1 = C1 + B1
            IEnumerable<string> set = spreadsheet.SetContentsOfCell("A1", "10.0");
            HashSet<string> testSet = new HashSet<string>(set);
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "A1", "D1", "E1" }));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For Formula
        /// </summary>
        [TestMethod]
        public void SetCellContentsSet1()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "5.0");
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("A1 + 7 / A2").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            IEnumerable<string> set = spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1").ToString());
            HashSet<string> testSet = new HashSet<string>(set);
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() {"A1", "E1","D1"}));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        public void SetCellContentsSet2()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "5.0");
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            IEnumerable<string> set = spreadsheet.SetContentsOfCell("A1", new Formula("C1").ToString());
            HashSet<string> testSet = new HashSet<string>(set);
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "A1", "D1"}));
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTest()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "5.0");
            spreadsheet.SetContentsOfCell("D1", "=" +  new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+A1").ToString()); // A1 = C1 + B1

        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTest2()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "=" + new Formula("A1").ToString());
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+E1").ToString()); // A1 = C1 + B1

        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTestDouble()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "=" + new Formula("A1").ToString());
            spreadsheet.SetContentsOfCell("A1", "34");
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+E1").ToString()); // A1 = C1 + B1

        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTestFormula()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "=" + new Formula("A1").ToString());
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("5+5").ToString());
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+E1").ToString()); // A1 = C1 + B1

        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTest3()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "=" + new Formula("A1").ToString());
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1").ToString()); // D1 = A1

        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// This tests the case where we test the sets that get based by set Cell contents of A1 should be C1 and B1
        /// For string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(CircularException))]
        public void CircularErrorTest4()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "=" + new Formula("C1").ToString());
        }


        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.GetCellContents("7A");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName2()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("7A", "text");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName3()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("& a + ase", "text");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName4()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("", "text");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case for formula
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName6()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("", "=" + new Formula("A1 + A2").ToString());
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case for double
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName5()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("", "15.0");
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void InvalidName7()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            String stuff = null;
            contents.SetContentsOfCell("A1", stuff);
        }

        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case for formula
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(FormulaFormatException))]
        public void InvalidName8()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            Formula test = null;
            contents.SetContentsOfCell("a1", "=" + test);
        }


        /// <summary>
        /// A test case that checks an average case for the get cell contents method
        /// Tests the null case for string
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void InvalidName9()
        {
            SS.Spreadsheet contents = new SS.Spreadsheet(); // Blank spreadsheet
            contents.SetContentsOfCell("+a1", "test");
        }

        /// <summary>
        /// Tests the active cells method
        /// </summary>
        [TestMethod]
        public void EmptyReturnTest()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "5.0");
            spreadsheet.SetContentsOfCell("D1", "=" + new Formula("A1").ToString()); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("C1+A1").ToString()); // A1 = C1 + B1
            IEnumerable<string> set = spreadsheet.SetContentsOfCell("A1", "10.0");
            HashSet<string> testSet = new HashSet<string>(spreadsheet.GetNamesOfAllNonemptyCells());
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "A1", "D1", "E1", "C1"}));
        }

        /// <summary>
        /// This Tests what happens if adding onto the same cell works with text
        /// </summary>
        [TestMethod]
        public void StringOverlapTest()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "test");
            spreadsheet.SetContentsOfCell("C1", "no"); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("C1+A1").ToString()); // A1 = C1 + B1
            HashSet<string> testSet = new HashSet<string>(spreadsheet.GetNamesOfAllNonemptyCells());
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "A1", "E1", "C1" }));
            Assert.IsTrue(spreadsheet.GetCellContents("C1").Equals("no"));
        }

        /// <summary>
        /// This Tests what happens if adding onto the same cell works with text is an invalid cell
        /// </summary>
        [TestMethod]
        public void StringOverlapTest2()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "test");
            spreadsheet.SetContentsOfCell("C1", "no"); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("C1+A1").ToString()); // A1 = C1 + B1
            HashSet<string> testSet = new HashSet<string>(spreadsheet.GetNamesOfAllNonemptyCells());
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "C1", "A1", "E1"}));
            Assert.IsTrue(spreadsheet.GetCellContents("C1").Equals("no"));
        }

        /// <summary>
        /// This Tests what happens if adding onto the same cell works with text is an invalid cell
        /// THIS TESTS EMPTY STRING ENTERED INTO E1
        /// </summary>
        [TestMethod]
        public void StringOverlapTest3()
        {
            SS.Spreadsheet spreadsheet = new SS.Spreadsheet(); // Blank spreadsheet
            spreadsheet.SetContentsOfCell("C1", "test");
            spreadsheet.SetContentsOfCell("C1", "1.0"); // D1 = A1
            spreadsheet.SetContentsOfCell("A1", "=" + new Formula("C1+B1").ToString()); // A1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "=" + new Formula("C1+A1").ToString()); // E1 = C1 + B1
            spreadsheet.SetContentsOfCell("E1", "");
            HashSet<string> testSet = new HashSet<string>(spreadsheet.GetNamesOfAllNonemptyCells());
            Assert.IsTrue(testSet.SetEquals(new HashSet<string>() { "C1", "A1"}));
        }





        /// <summary>
        /// Let us start testing the save functiosn for the spreadsheet
        /// </summary>
        [TestMethod]
        public void SaveTest()
        {
            Spreadsheet save = new Spreadsheet();
            save.SetContentsOfCell("A1", "53");
            save.SetContentsOfCell("A2", "five");
            save.SetContentsOfCell("A2", "72");
            save.SetContentsOfCell("A3", "=A1+A2");
            save.Save("test1.xml");
        }


        /// <summary>
        /// TESTS THE VERSION CONTROL
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void LoadTest3()
        {
            Spreadsheet save = new Spreadsheet();
            save.SetContentsOfCell("A1", "53");
            save.SetContentsOfCell("A2", "five");
            save.SetContentsOfCell("A2", "72");
            save.SetContentsOfCell("A3", "=A1+A2");
            save.Save("test2.xml");
            Spreadsheet load = new Spreadsheet("test2.xml", s => true, s => s, "wrong version");

        }



        /// <summary>
        /// Let us start testing the save functiosn for the spreadsheet
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest()
        {
            Spreadsheet save = new Spreadsheet();
            save.SetContentsOfCell("_A1", "53");
        }

        /// <summary>
        /// Let us start testing the save functiosn for the spreadsheet
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void ErrorTest4()
        {
            Spreadsheet save = new Spreadsheet("RandomHtml", s => true, s => s, "default");
        }

        /// <summary>
        /// TESTS GET CELL CONTEST
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest6()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.GetCellValue("C1");
            getCell.GetCellValue("_C1");
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest7()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.GetCellContents("C1");
            getCell.GetCellContents("_C1");
        }
        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(FormulaFormatException))]
        public void ErrorTest10()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.SetContentsOfCell("", "=");
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest11()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.SetContentsOfCell("", "=A1");
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest12()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.SetContentsOfCell("", "A1");
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest13()
        {
            Spreadsheet getCell = new Spreadsheet();
            getCell.SetContentsOfCell("", "12");
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        //[ExpectedException(typeof(ArgumentException))]
        public void ErrorTest14()
        {
            Spreadsheet getCell = new Spreadsheet();
            //getCell.SetContentsOfCell("A1", null);
        }

        /// <summary>
        /// TESTS MORE ERRORS
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void ErrorTest9()
        {
            Spreadsheet save = new Spreadsheet("tag.xml", s => true, s => s, "default");
        }

        /// <summary>
        /// TESTS MORE ERRORS OF ISVALID
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest15()
        {
            Spreadsheet save = new Spreadsheet( s => false, s => s, "default");
            save.SetContentsOfCell("A1", "7"); // Should be invalid
        }

        /// <summary>
        /// TESTS MORE ERRORS OF ISVALID
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest16()
        {
            Spreadsheet save = new Spreadsheet(s => false, s => s, "default");
            save.GetCellContents("A1"); // Should be invalid
        }
        /// <summary>
        /// TESTS MORE ERRORS OF ISVALID
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest17()
        {
            Spreadsheet save = new Spreadsheet(s => false, s => s, "default");
            save.GetCellValue("A1"); // Should be invalid
        }
        /// <summary>
        /// TESTS MORE ERRORS OF ISVALID
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void ErrorTest18()
        {
            Spreadsheet save = new Spreadsheet(s => false, s => s, "default");
            save.SetContentsOfCell("74", "5"); // Should be invalid
        }
    }
    }


