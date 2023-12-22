

/*
using System;

namespace MSerialNo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }
}
*/
using System;
using System.Linq;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace MSerialNo
{
    class Program
    {

        //Declaration of variable
        static int startSerialNo = 0;
        static int rangeSerialNo = 0;
        static decimal BoxNos = 0;
        static int remainingPics = 0;

        static void Main(string[] args)
        {


            //calling the Main Manu
            MainMenu();



            /*
            // Create a list of accounts.
            var bankAccounts = new List<Account>
            {
                new Account {
                              ID = 345678,
                              Balance = 541.27
                            },
                new Account {
                              ID = 1230221,
                              Balance = -127.44
                            }
            };
            */

            // Display the list in an Excel spreadsheet.
            // DisplayInExcel(bankAccounts);
            // ExportInnerBoxToExcel(231060477,50);

            // Create a Word document that contains an icon that links to
            // the spreadsheet.
            // CreateIconInWordDoc();
        }


        /// <summary>
        /// *********************************************
        /// Method for main menu
        /// Main Manu
        /// ********************************************
        /// </summary>
        static void MainMenu()
        {
            Console.WriteLine("Meter Serial Numbers ");
            Console.WriteLine("===================== ");
            Console.WriteLine("");
            Console.WriteLine("MAIN MENU ");
            Console.WriteLine("==========");
            Console.WriteLine(" 1. Enter [S]erial Number ");
            Console.WriteLine(" 2. [P]review");
            Console.WriteLine(" 3. [E]xport");
            Console.WriteLine(" 4. [H]elp");
            Console.WriteLine(" 5. E[x]it");
            Console.WriteLine();
            Console.Write("Enter Choice :");
            //P221046002
            string ReadMenu = Console.ReadLine();

            if (ReadMenu.Equals("1") || ReadMenu.Equals("s") || ReadMenu.Equals("S"))
            {
                //Entering the range of serial numbers that can be accessed any where in the system
                Console.Clear();
                Console.WriteLine("SERIAL NUMBER(S) MENU ");
                Console.WriteLine("=====================");
                Console.WriteLine("Press 'B' and Press Enter To Main Manu");
                Console.WriteLine("");
                string getSerial = "";


                int countString = 0;

                //Loop through to count user entery
               for(; ; )
                {
                    //Get serial numbers
                    Console.Write("Enter Serial Number (9 Digits):");
                    getSerial = Console.ReadLine();
                    countString = getSerial.Count();

                   if (countString == 9 || countString == 8)
                    {
                        break;
                    }

                    //Removing the prefix p 
                    //  if (countString == 10 && getSerial[0].Equals("P") || getSerial[0].Equals("P"))
                    //{

                    //  }


                }
               // while (countString != 9 || countString != 8);


                //Check user entry again if is digits
                try
                { startSerialNo = Convert.ToInt32(getSerial); }
                catch (Exception e)
                {
                    Console.Clear();
                    Console.Beep();
                    Console.WriteLine("");
                    Console.WriteLine("Sorry! Digits Only " + e.Message);
                    Console.WriteLine("");
                    MainMenu();
                }
                bool recheck = false;
                //Loop through to check range of numbers
                do
                {
                    //Get serial numbers
                    Console.Write("Enter Range of Numbers:");
                    getSerial = Console.ReadLine();

                    //Check user entry again if is digits
                    try
                    {
                        recheck = true;
                        rangeSerialNo = Convert.ToInt32(getSerial);

                    }
                    catch (Exception e)
                    {
                        recheck = false;
                    }
                }
                while (recheck == false);


                //Display user entry and total range
                int tSum = (startSerialNo + rangeSerialNo - 1);

                //Get the number of boxes
                BoxNos = rangeSerialNo/8;
                remainingPics = rangeSerialNo % 8;
                Console.Clear();
                Console.WriteLine("");
                Console.WriteLine(" The Serial Number Starts From :" + startSerialNo + " To :" + tSum);
                if (remainingPics==0)
                {
                    Console.WriteLine(" The Number of Boxes : " + BoxNos);
                }
                else
                {
                    Console.WriteLine(" The Number of Boxes : " + BoxNos + " The Last Box Has " + remainingPics + " Pics");
                }
                
                Console.WriteLine("");
                MainMenu();

            }
            else if (ReadMenu.Equals("2") || ReadMenu.Equals("p") || ReadMenu.Equals("P"))
            {
                //Showing the preview manu

                //Check if serial numbers are entered
                //before user can get access to preview
                if (startSerialNo == 0 && rangeSerialNo == 0)
                {
                    Console.Clear();
                    Console.Beep();
                    Console.WriteLine("");
                    Console.WriteLine("Sorry! Serial Number and Range of numbers are needed");
                    Console.WriteLine("");
                    MainMenu();

                }
                else
                {
                    // PreviewList(); 
                    PreviewSerialNos(startSerialNo, rangeSerialNo);
                   // ExportInnerBoxToExcel(startSerialNo, rangeSerialNo);
                }

            }
            else if (ReadMenu.Equals("3") || ReadMenu.Equals("e") || ReadMenu.Equals("E"))
            {
                //Show the export menu
                //Check if serial numbers are entered
                //before user can get access to preview
                if (startSerialNo == 0 && rangeSerialNo == 0)
                {
                    Console.Clear();
                    Console.Beep();
                    Console.WriteLine("");
                    Console.WriteLine("Sorry! Serial Number and Range of numbers are needed");
                    Console.WriteLine("");
                    MainMenu();

                }
                else
                {

                    ExportList();
                }
            }
            else if (ReadMenu.Equals("4") || ReadMenu.Equals("h") || ReadMenu.Equals("H"))
            {
                //Showing Help menu
                // HelpManu();

                Console.Clear();
                Console.Beep();
                Console.WriteLine("");
                Console.WriteLine("Sorry! Still Under Construction");
                Console.WriteLine("");
                MainMenu();
            }
            else if (ReadMenu.Equals("5") || ReadMenu.Equals("x") || ReadMenu.Equals("X"))
            {
                //Closing the entire application
                Console.Clear();
                Console.WriteLine("\n\n\n\n\n\n\n\n");
                Console.WriteLine("\t\t\t**************************************************");
                Console.WriteLine("\t\t\t       THANK YOU FOR USING THIS SOFTWARE          ");
                Console.WriteLine("\t\t\t**************************************************");
                Console.WriteLine("\n\n\n");
                System.Environment.Exit(0);
            }
            else
            {
                Console.Clear();
                Console.Beep();
                Console.WriteLine("");
                Console.WriteLine("Invalid Choice, Try again");
                Console.WriteLine("");
                MainMenu();

            }

            //Console.ReadLine()
        }

        /// <summary>
        /// *********************************************
        /// Method for Preview Menu
        /// Preview Manu
        /// ********************************************
        /// </summary>
        /// 
        static void PreviewList()
        {
            Console.WriteLine("PREVIEW MENU ");
            Console.WriteLine("============");
            Console.WriteLine(" 1. view List of [S]erial Numbers");
        }

        /// <summary>
        /// *********************************************
        /// Method for Export Menu
        /// Export Manu
        /// ********************************************
        /// </summary>
        /// 
        static void ExportList()
        {
            Console.Clear();
            Console.WriteLine("Export MENU ");
            Console.WriteLine("============");
            Console.WriteLine(" 1. [S]erial Numbers");
            Console.WriteLine(" 2. [I]nner Box");
            Console.WriteLine(" 3. [O]uter Box");
            Console.WriteLine(" 4. [B]ack To Main Manu");

            //Using Infinit loop 
            for(; ; )
            {
                Console.Write("Choose One :");
                string exportToExcel = Console.ReadLine();

                if (exportToExcel.Equals("1") || exportToExcel.Equals("s") || exportToExcel.Equals("S"))
                {
                    //Export Serial Numbers to excel
                    //by calling the method
                    ExportListToExcel(startSerialNo, rangeSerialNo);
                }
                else if (exportToExcel.Equals("2") || exportToExcel.Equals("i") || exportToExcel.Equals("I"))
                {
                    //Export Inner Box to excel
                    ExportInnerBoxToExcel(startSerialNo, rangeSerialNo);
                }

                else if (exportToExcel.Equals("3") || exportToExcel.Equals("o") || exportToExcel.Equals("O"))
                {
                    //Export Inner Box to excel
                    // ExportListToExcel(startSerialNo, rangeSerialNo);
                    //   CheckPointHere(startSerialNo, rangeSerialNo);
                    ExportOuterBoxToExcel(startSerialNo, rangeSerialNo);

                }
                else if (exportToExcel.Equals("4") || exportToExcel.Equals("t") || exportToExcel.Equals("t"))
                {
                    //Export Inner Box to excel
                    // ExportListToExcel(startSerialNo, rangeSerialNo);
                    //   CheckPointHere(startSerialNo, rangeSerialNo);
                    ExportThreeOuterBoxToExcel(startSerialNo, rangeSerialNo);

                }
                else if (exportToExcel.Equals("5") || exportToExcel.Equals("b") || exportToExcel.Equals("B"))
                {
                    Console.Clear();
                    Console.WriteLine("");
                    MainMenu();
                }
                else
                {
                    Console.Write("Invalid entry, try again!");
                }
            }
        }


        /// <summary>
        /// *********************************************
        /// Method for Help Menu
        /// Help Manu
        /// ********************************************
        /// </summary>
        /// 
        static void HelpManu()
        {
            Console.WriteLine("MAIN MENU ");
            Console.WriteLine("==========");
            Console.WriteLine(" 1. Enter [S]erial Number ");
            Console.WriteLine(" 2. [P]review");
            Console.WriteLine(" 3. [E]xport");
            Console.WriteLine(" 4. [H]elp");
            Console.WriteLine(" 5. E[x]it");
            Console.WriteLine();
        }
        //Method to preview
        static void PreviewSerialNos(int sPoint, int ePoint)
        {
            Console.WriteLine("List of " + ePoint+" Serial Numbers");
            Console.WriteLine("===================================");
            for (int i = 0; i < ePoint; i++)
            {
                int tSum = (sPoint + i);
                Console.WriteLine("P" + tSum);
            }

            for(; ; )
            {
                Console.WriteLine("");
                Console.Write("To return to Menu Press 'B' :");
                
                 string KeyPrssing = Console.ReadLine();

                if(KeyPrssing.Equals("B") || KeyPrssing.Equals("b"))
                {
                    Console.Clear();
                    Console.WriteLine("");
                    MainMenu();
                }

            }

            
         


        }


        /// <summary>
        /// *********************************************
        /// Method for Export Menu
        /// List of Serial Numbers
        /// ********************************************
        /// </summary>
        /// 
        static void ExportListToExcel(int StartValue, int RangeValue, string MType = "Single")
        {
            //Declaring and initializing variable
            int i = 0;
            string[] serialNos= new string[RangeValue];
            int rowNo = 1;

            //Creating an instance of an excel
            var exportEx = new Excel.Application();

            //Show the object
           // exportEx.Visible = true;

            //Add workbook
            exportEx.Workbooks.Add();

            //Calling single worksheet
            Excel._Worksheet wkSheet = exportEx.ActiveSheet;

            //Using style or fonts
            wkSheet.Cells[rowNo, "A"].Style.Font.size = 24;
            wkSheet.Cells[rowNo, "A"].Style.Font.Bold = true;

            //It prints to excel if the row is equal to 1 or A1
            wkSheet.Cells[rowNo, "A"] =  MType + " Phase";
            for (i = 0; i < RangeValue; i++)
            {
                // Console.WriteLine(StartValue+i);
                serialNos[i] = "P" + (StartValue + i).ToString();
                //  Console.WriteLine(serialNos[i]);
                rowNo++;
                wkSheet.Cells[rowNo, "A"] = serialNos[i]; //'0x800AC472

                //Using style or fonts
                wkSheet.Cells[rowNo, "A"].Style.Font.size = 20;
                wkSheet.Cells[rowNo, "B"].Style.Font.Bold = true;

            }
           
            wkSheet.Columns[1].AutoFit();

            //Show the object
            exportEx.Visible = true;

       

            Console.WriteLine("Successfully Exported");

        }

        //Check Point
        static void CheckPointHere(int StartValue, int RangeValue, string MType = "Single")
        {
            Console.WriteLine("Meter Serial No.");
           int countCheck = 0;
            for (int i = 0; i < RangeValue; i++)
            {
                countCheck++;
                int tSum = (StartValue + i);
                Console.WriteLine("P" + tSum);

                if(countCheck==8)
                {
                    Console.WriteLine("Meter Serial No.");
                    countCheck = 0;
                }


            }

        }

        /// <summary>
        /// *********************************************
        /// Method for Export Menu
        /// List of Serial Numbers for Inner Box
        /// ********************************************
        /// </summary>
        /// 
        static void ExportInnerBoxToExcel(int StartValue, int RangeValue, string MType = "Single")
        {
            //Declaring and initializing variable
            int i = 0;
            string[] serialNos = new string[RangeValue];
            int rowNo = 1;
            int rowNo2 = 2;
            int rowNo3 = 3;

            //Creating an instance of an excel
            var exportEx = new Excel.Application();


            //Add workbook
            exportEx.Workbooks.Add();

            //Calling single worksheet
            Excel._Worksheet wkSheet = exportEx.ActiveSheet;

           

            //Loop to generate Inner Box Serial Numbers
            for (i = 0; i < RangeValue; i++)
            {
                
                //Condition to check each position in a cell
                if (i == 0)
                {
                    wkSheet.Cells[rowNo, "A"] = "Meter Type";
                    wkSheet.Cells[rowNo, "B"] = MType + " Phase";

                    wkSheet.Cells[rowNo2, "A"] = "Meter Serial No.";
                    wkSheet.Cells[rowNo2, "B"] = serialNos[i] = "P" + (StartValue + i).ToString();

                    wkSheet.Cells[rowNo3, "A"] = "";
                    wkSheet.Cells[rowNo3, "B"] = "";
                }
                else
                {
                    //Adding 3 to each counter
                    rowNo += 3;
                    rowNo2 += 3;
                    rowNo3 += 3;

                    wkSheet.Cells[rowNo, "A"] = "Meter Type";
                    wkSheet.Cells[rowNo, "B"] = MType + " Phase";

                    wkSheet.Cells[rowNo2, "A"] = "Meter Serial No.";
                    wkSheet.Cells[rowNo2, "B"] = serialNos[i] = "P" + (StartValue + i).ToString();


                    wkSheet.Cells[rowNo3, "A"] = "";
                    wkSheet.Cells[rowNo3, "B"] = "";


                }


                //Using style or fonts
                wkSheet.Cells[rowNo, "A"].Style.Font.size = 20;
                wkSheet.Cells[rowNo, "B"].Style.Font.size = 20;
                wkSheet.Cells[rowNo, "A"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo, "B"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo, "A"].Style.Font.Bold = true;
                wkSheet.Cells[rowNo, "B"].Style.Font.Bold = true;

                wkSheet.Cells[rowNo2, "A"].Style.Font.size = 20;
                wkSheet.Cells[rowNo2, "B"].Style.Font.size = 20;
                wkSheet.Cells[rowNo2, "A"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo2, "B"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo2, "A"].Style.Font.Bold = true;
                wkSheet.Cells[rowNo2, "B"].Style.Font.Bold = true;

                wkSheet.Cells[rowNo3, "A"].Style.Font.size = 20;
                wkSheet.Cells[rowNo3, "B"].Style.Font.size = 20;
                wkSheet.Cells[rowNo3, "A"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo3, "B"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[rowNo3, "A"].Style.Font.Bold = true;
                wkSheet.Cells[rowNo3, "B"].Style.Font.Bold = true;

            }

            //Auto fit each cell
            wkSheet.Columns[1].AutoFit();
            wkSheet.Columns[2].AutoFit();

            //Show the object
            exportEx.Visible = true;


            //message
            Console.WriteLine("Successfully Exported");

        }



        /// <summary>
        /// *********************************************
        /// Method for Export Menu
        /// List of Serial Numbers for Outer Box
        /// ********************************************
        /// </summary>
        /// 
        static void ExportOuterBoxToExcel(int StartValue, int RangeValue, string MType = "Single")
        {
            //Declaring and initializing variable
           // int i = 0;
            string[] serialNos = new string[RangeValue];
            int rowNo = 1;
            int rowNo2 = 2;
            int rowNo3 = 3;

            //Range divided by 2
            int RangeDiv = RangeValue / 2;

            //Set cell characters
            string SetA = "A";
            string SetB = "B";

            DateTime startTimer= new DateTime();
            DateTime EndTimer = new DateTime();
            

            //Creating an instance of an excel
            var exportEx = new Excel.Application();


            //Add workbook
            exportEx.Workbooks.Add();

            //Calling single worksheet
            Excel._Worksheet wkSheet = exportEx.ActiveSheet;


            //Generating Serial Number in B Cell
            int browNo = 1;
            wkSheet.Cells[browNo, SetA].Style.Font.size = 24;
            wkSheet.Cells[browNo, SetB].Style.Font.size = 24;
            wkSheet.Cells[browNo, SetA] ="Box No.";
            wkSheet.Cells[browNo, SetB] = "Meter Serial No.";
            int countCheck = 0;
            for (int i = 0; i < RangeValue; i++)
            {
                wkSheet.Cells[browNo, SetA].Style.Font.size = 20;
                wkSheet.Cells[browNo, SetB].Style.Font.size = 20;
                browNo++;
                countCheck++;
                int tSum = (StartValue + i);
                wkSheet.Cells[browNo, SetB] = "P" + tSum;

                if (countCheck == 9)
                {
                    wkSheet.Cells[browNo, SetA].Style.Font.size = 24;
                    wkSheet.Cells[browNo, SetB].Style.Font.size = 24;
                    wkSheet.Cells[browNo, SetA] = "Box No.";
                    wkSheet.Cells[browNo, SetB] = "Meter Serial No.";
                    countCheck = 0; 
                    
                   
                    i--;
                }
                wkSheet.Cells[browNo, SetA].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[browNo, SetB].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[browNo, SetA].Style.Font.Bold = true;
                wkSheet.Cells[browNo, SetB].Style.Font.Bold = true;

                if (i== RangeDiv)
                {
                    //Reset cell characters
                    SetA = "D";
                    SetB = "E";
                    browNo = 1;
                }


            }

       

            //Auto fit each cell
            wkSheet.Columns[1].AutoFit();
            wkSheet.Columns[2].AutoFit();

            //Show the object
            exportEx.Visible = true;



            //message
            Console.WriteLine("Successfully Exported");

        }

        /// <summary>
        /// *********************************************
        /// Method for Export Menu
        /// List of Serial Numbers for 3 Phase Outer Box
        /// ********************************************
        /// </summary>
        /// 
        static void ExportThreeOuterBoxToExcel(int StartValue, int RangeValue, string MType = "Three")
        {
            //Declaring and initializing variable
            // int i = 0;
            string[] serialNos = new string[RangeValue];
            int rowNo = 1;
            int rowNo2 = 2;
            int rowNo3 = 3;

            //Creating an instance of an excel
            var exportEx = new Excel.Application();


            //Add workbook
            exportEx.Workbooks.Add();

            //Calling single worksheet
            Excel._Worksheet wkSheet = exportEx.ActiveSheet;


            //Generating Serial Number in B Cell
            int browNo = 1;
            wkSheet.Cells[browNo, "A"].Style.Font.size = 24;
            wkSheet.Cells[browNo, "B"].Style.Font.size = 24;
            wkSheet.Cells[browNo, "A"] = "Box No.";
            wkSheet.Cells[browNo, "B"] = "Meter Serial No.";
            int countCheck = 0;
            for (int i = 0; i < RangeValue; i++)
            {
                wkSheet.Cells[browNo, "A"].Style.Font.size = 20;
                wkSheet.Cells[browNo, "B"].Style.Font.size = 20;
                browNo++;
                countCheck++;
                int tSum = (StartValue + i);
                wkSheet.Cells[browNo, "B"] = "P" + tSum;

                if (countCheck == 5)
                {
                    wkSheet.Cells[browNo, "A"].Style.Font.size = 24;
                    wkSheet.Cells[browNo, "B"].Style.Font.size = 24;
                    wkSheet.Cells[browNo, "A"] = "Box No.";
                    wkSheet.Cells[browNo, "B"] = "Meter Serial No.";
                    countCheck = 0;


                    i--;
                }
                wkSheet.Cells[browNo, "A"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[browNo, "B"].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wkSheet.Cells[browNo, "A"].Style.Font.Bold = true;
                wkSheet.Cells[browNo, "B"].Style.Font.Bold = true;
            }



            //Auto fit each cell
            wkSheet.Columns[1].AutoFit();
            wkSheet.Columns[2].AutoFit();

            //Show the object
            exportEx.Visible = true;


            //message
            Console.WriteLine("Successfully Exported");

        }


    }
   
}