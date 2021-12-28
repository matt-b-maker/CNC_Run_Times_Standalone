using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using EPDM.Interop.epdm;
using System.Linq;
using System.Threading;

namespace CNC_Run_Times_and_Material_Counts
{
    class Program
    {
        static void Main(string[] args)
        {
            bool exceptionEncountered = false;

            string command;
            string prodNum;
            bool isOrganizedProperly; 

            IEdmVault21 CurrentVault = new EdmVault5() as IEdmVault21;
            IEdmSearch9 _search;
            IEdmSearchResult5 _searchResult;

            float prodRuntime = 0f;
            float totalRunTime = 0f;

            string fullPath;
            string cncPath;
            string filePath;

            List<string[]> files = new List<string[]>();
            List<RunTime> runTimeObjects = new List<RunTime>();
            List<string> filePaths = new List<string>();
            List<Material> materials = new List<Material>();

            string message = "Welcome to the virtual hot dog creator. Time to create some virtual hot dogs.";

            WriteMessage(message);

            Thread.Sleep(2000);

            WriteMessage("\nLogging into the PDM...");

            try
            {
                CurrentVault.LoginAuto("CreativeWorks", 0);
            }
            catch
            {
                WriteMessage("You need to be logged into the PDM, genius.");
                Console.ReadKey();
                Console.Clear();
            }

            WriteMessage("\nLogged into the PDM successfully\n");
            Thread.Sleep(1000);

            do
            {
                Console.WriteLine("\nEnter the four numbers of the PROD # or type \"STOP\" to stop: ");
                command = Console.ReadLine();

                if (CheckCommand(command))
                {
                    prodNum = command;
                }
                else
                {
                    break;
                }

                _search = (IEdmSearch9)CurrentVault.CreateSearch2();
                _search.FindFiles = true;

                _search.Clear();
                _search.StartFolderID = CurrentVault.GetFolderFromPath("C:\\CreativeWorks").ID;

                _search.FileName = "PROD-" + prodNum + ".sldasm";

                _search.GetFirstResult();

                if (exceptionEncountered)
                {
                    _searchResult = null;
                }
                else
                {
                    _searchResult = _search.GetFirstResult();
                }

                if (_searchResult != null)
                {
                    cncPath = GetCNCPath(_searchResult.Path);
                    _search.Clear();
                    _search.StartFolderID = CurrentVault.GetFolderFromPath(cncPath).ID;
                    _search.FileName = "*.sldasm";
                    _search.GetFirstResult();

                    if (exceptionEncountered)
                    {
                        _searchResult = null;
                    }
                    else
                    {
                        _searchResult = _search.GetFirstResult();
                    }

                    if (_searchResult != null)
                    {
                        isOrganizedProperly = true;
                    }
                    else
                    {
                        isOrganizedProperly = false; 
                    }
                }
                else
                {
                    Console.WriteLine("Didn't find anything for that");
                    continue;
                }

                //Path 1: Not organized correctly, add file names to material objects
                if (!isOrganizedProperly)
                {
                    Console.WriteLine("The folder in question either doesn't exist or isn't organized properly.");
                    Console.WriteLine("Press enter to continue searching for shopbot files or type \"STOP\" to end program.");
                    command = Console.ReadLine();

                    if (CheckCommand(command))
                    {
                        //Search for ShopBot files only using the PROD#. Things could get messy here
                        Console.WriteLine("Double check that this path didn't return any files that shouldn't be in the list.");
                        Console.WriteLine("If the folder wasn't organized properly, it will be difficult to determine whether this program");
                        Console.WriteLine("is returning the correct information.\n");

                        _search.FileName = "*.sbp";
                        _search.GetFirstResult();

                        if (exceptionEncountered)
                        {
                            _searchResult = null;
                        }
                        else
                        {
                            _searchResult = _search.GetFirstResult();
                        }

                        if (_searchResult != null)
                            filePath = _searchResult.Path;
                        else
                        {
                            Console.WriteLine("Didn't find anything. Try again.");
                            continue;
                        }

                        AddMaterial(false, materials, filePath);

                        filePaths.Add(filePath);
                        try
                        {
                            files.Add(File.ReadAllLines(filePath));
                        }
                        catch
                        {
                            Console.WriteLine("There were either no shopbot files or you need to copy the files to get local copies for this program to read. Try again.");
                            break; 
                        }

                        while (_searchResult != null)
                        {
                            _searchResult = _search.GetNextResult();

                            if (_searchResult != null)
                            {
                                filePath = _searchResult.Path;

                                AddMaterial(false, materials, filePath);

                                filePaths.Add(filePath);
                                try
                                {
                                    files.Add(File.ReadAllLines(filePath));
                                }
                                catch
                                {
                                    Console.WriteLine("There were either no shopbot files or you need to copy the files to get local copies for this program to read. Try again.");
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (files.Count != 0)
                        {
                            runTimeObjects = CalculateRunTime(files, runTimeObjects, filePaths);

                            runTimeObjects = runTimeObjects.OrderBy(o => o.FileName).ToList();

                            foreach (var runTime in runTimeObjects)
                            {
                                Console.WriteLine(runTime.FileName + "\n" + runTime.Seconds);
                                prodRuntime += runTime.Seconds;
                            }

                            totalRunTime += prodRuntime;

                            Console.WriteLine($"\nPROD run time in hours: {Math.Round((prodRuntime / 3600), 2)}");
                            Console.WriteLine($"Total run time in hours: {Math.Round((totalRunTime / 3600), 2)}");

                            prodRuntime = 0;

                            runTimeObjects.Clear();
                            filePaths.Clear();
                            files.Clear();
                            _searchResult = null;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }

                //Path 2: Organized, check material for duplicates and make totals
                else
                {
                    Console.WriteLine("Found a properly organized folder. Finding ShopBot files now.\n");
                    _search.Clear();
                    _search.StartFolderID = CurrentVault.GetFolderFromPath(cncPath + "Programs\\").ID;
                    _search.FileName = "*.sbp";
                    _search.GetFirstResult();

                    if (exceptionEncountered)
                    {
                        _searchResult = null;
                    }
                    else
                    {
                        _searchResult = _search.GetFirstResult();
                    }

                    if (_searchResult != null)
                        filePath = _searchResult.Path;
                    else
                    {
                        Console.WriteLine("Didn't find anything. Try again.");
                        continue;
                    }

                    if (_searchResult != null && !_searchResult.Path.ToUpper().Contains("PARTS") && !_searchResult.Path.ToUpper().Contains("RECUTS"))
                    {
                        filePaths.Add(filePath);
                        if (!CheckForTwoParter(filePath))
                            AddMaterial(true, materials, filePath);
                        try
                        {
                            files.Add(File.ReadAllLines(filePath));
                        }
                        catch
                        {
                            Console.WriteLine("There were either no shopbot files or you need to copy the files to get local copies for this program to read. Try again.");
                            break;
                        }
                    }

                    while (_searchResult != null)
                    {
                        _searchResult = _search.GetNextResult();

                        if (_searchResult != null && !_searchResult.Path.ToUpper().Contains("PARTS") && !_searchResult.Path.ToUpper().Contains("RECUTS"))
                        {
                            filePath = _searchResult.Path;
                            filePaths.Add(filePath);
                            if (!CheckForTwoParter(filePath))
                                AddMaterial(true, materials, filePath);
                            try
                            {
                                files.Add(File.ReadAllLines(filePath));
                            }
                            catch
                            {
                                Console.WriteLine("There were either no shopbot files or you need to copy the files\n to get local copies for this program to read. Try again.");
                                break;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }

                    if (files.Count != 0)
                    {
                        runTimeObjects = CalculateRunTime(files, runTimeObjects, filePaths);

                        runTimeObjects = runTimeObjects.OrderBy(o => o.FileName).ToList();

                        foreach (var runTime in runTimeObjects)
                        {
                            Console.WriteLine(runTime.FileName + "\n" + runTime.Seconds);
                            prodRuntime += runTime.Seconds;
                        }

                        totalRunTime += prodRuntime;

                        Console.WriteLine($"\nPROD run time: {Math.Round((prodRuntime / 3600), 2)}");
                        Console.WriteLine($"Total Run Time: {Math.Round((totalRunTime / 3600), 2)}");

                        prodRuntime = 0;

                        runTimeObjects.Clear();
                        filePaths.Clear();
                        files.Clear();
                        _searchResult = null;
                    }
                }


            } while (CheckCommand(command));

            if (totalRunTime > 0)
            {
                Console.Clear();
                message = "Thank you for using the Creative Works Virtual Hot Dog Creator.\nPress enter to view your hot dogs:";

                WriteMessage(message);

                Console.ReadKey();

                Console.Clear();

                Random letter = new();

                /*

                for (int i = 5; i > 0; i--)
                {
                    Console.Clear();
                    Console.Write("Beginning complex calculation in {0}", i);
                    Thread.Sleep(1000);
                }

                /*
                for (int i = 0; i < 500; i++)
                {
                    for (int j = 0; j < Console.BufferWidth; j++)
                    {
                        Console.Write((char)letter.Next(65, 122));
                    }
                }
                Console.WriteLine("\n\nThis is a very complex calculation. Hang on, switching to green mode...");
                Thread.Sleep(4000);
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                for (int i = 0; i < 1000; i++)
                {
                    for (int j = 0; j < Console.BufferWidth; j++)
                    {
                        Console.Write((char)letter.Next(65, 122));
                    }
                }
                Console.WriteLine("\n\nCalculation Complete. Compiling information...");
                Thread.Sleep(5000);
                */
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"\nThe total run time of all files is {Math.Round((totalRunTime / 3600), 2)} hour(s)");
                for (int i = 0; i < materials.Count; i++)
                {
                    if (materials[i].Quantity == 0)
                    {
                        materials.Add(materials[i]);
                        materials.RemoveAt(i);
                    }
                }
                foreach(var material in materials)
                {
                    Console.WriteLine(material.Quantity == 0 ? material.Name : WriteMaterialInfo(material));
                }

                Console.WriteLine("\nWe're done here. Press enter to exit.");
                Console.ReadKey();
            }
            else
            {
                Console.Clear();
                Console.WriteLine("You didn't create any hot dogs this time. Shame.");
                Console.ReadKey();
            }
        }

        private static void WriteMessage(string message)
        {
            foreach (var letter in message)
            {
                if (letter != '\n' && letter != '\0')
                {
                    Console.Write(letter);
                    Thread.Sleep(20);
                }
                else
                {
                    Console.Write(letter);
                }
            }
        }

        private static string WriteMaterialInfo(Material material)
        {
            string info = material.Name;
            for (int i = 0; i < 30 - material.Name.Length; i++)
            {
                info += " ";
            }
            info += material.Quantity;
            return info;
        }

        private static bool CheckForTwoParter(string filePath)
        {
            //Sample file path where the function is looking for the number 2 at the end of the file name:
            //"C:\CreativeWorks\Designs\LT\CEN\PROD-0709\2-CNC\Programs\1-MDX 50\4_1-MDX 502.sbp"
            if (Convert.ToInt32(filePath[filePath.Length - 5].ToString()) >= 2)
                return true;
            return false;
        }

        private static void AddMaterial(bool isOrganized, List<Material> materials, string filePath)
        {
            if (isOrganized)
            {
                string currentName = GetMaterialName(filePath);
                if (materials.Count == 0)
                {
                    materials.Add(new Material(currentName, 1));
                }
                else
                {
                    foreach (var material in materials)
                    {
                        if (material.Name == currentName)
                        {
                            material.Quantity++;
                            return;
                        }
                    }
                    materials.Add(new Material(currentName, 1));
                }
            }
            else
            {
                Regex fileNamePattern = new Regex(@"C:\\.+\\");
                string name = fileNamePattern.Replace(filePath, "");
                materials.Add(new Material(name));
            }
        }

        private static string GetMaterialName(string filePath)
        {
            string[] cutPath = filePath.Split('\\');
            return cutPath[^2];
        }

        private static string GetCNCPath(string path)
        {
            Regex cncPattern = new(@"1-.+");

            string cncPath = cncPattern.Replace(path, "2-CNC\\");

            return cncPath;
        }

        private static bool CheckCommand(string command)
        {
            if (command.ToUpper().Contains('S') || command.ToUpper().Contains("CUCK"))
            {
                return false;
            }
            return true;
        }

        private static List<RunTime> CalculateRunTime(List<string[]> files, List<RunTime> runTimeObjects, List<string> filePaths)
        {
            //Point and move speed variables
            double moveSpeedXY = 0, moveSpeedZ = 0;
            /* jogSpeed will change once user input is accepted in the program*/
            double jogSpeed = 10;
            //These variable will represent the coordinates of the tool's previous position
            double x2 = 0, y2 = 0, z2 = 0;
            //This variable is to ensure that the first function in the conditional below will only happen once
            var zCount = 0;

            //Calculation Variables
            double runTime = 0;
            double totRunTime = 0;

            //Used to separate the strings
            char[] separators = new char[] { ',', ' ' };
            int nameCount = 0;

            //Go through each file
            foreach (string[] file in files)
            {
                //x and y will always start at zero, so at the beginning of each iteration (each file), they will be
                //set to zero
                double x1 = 0;
                double y1 = 0;

                //The beginning of shopbot files will always set the safe z height. Once the iterations find the 
                //line with safez in it, it will assign the value on the line to this variable
                double z1 = 0;
                foreach (string line in file)
                {
                    double distance;
                    if (line.StartsWith("MS"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        moveSpeedXY = Convert.ToDouble(subs[1]);
                        moveSpeedZ = Convert.ToDouble(subs[2]);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]}");
                    }
                    else if (line.StartsWith("JZ"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        if (zCount == 0)
                        {
                            z1 = Convert.ToDouble(subs[1]);
                            zCount++;
                        }
                        else
                        {
                            z2 = Convert.ToDouble(subs[1]);
                            if (z1 == z2)
                            {
                                continue;
                            }
                            else if (z1 > z2)
                            {
                                distance = Math.Abs(z1 - z2);
                                runTime += GetTime(distance, moveSpeedZ);
                                z1 = z2;
                            }
                            else if (z2 > z1)
                            {
                                distance = Math.Abs(z2 - z1);
                                runTime += GetTime(distance, moveSpeedZ);
                                z1 = z2;
                            }
                        }
                    }
                    else if (line.StartsWith("M3"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]} {subs[3]}");
                        x2 = Convert.ToDouble(subs[1]);
                        y2 = Convert.ToDouble(subs[2]);
                        z2 = Convert.ToDouble(subs[3]);
                        if (IsZ(x1, y1, x2, y2))
                        {
                            distance = Math.Abs(z2 - z1);
                            runTime += GetTime(distance, moveSpeedZ);
                        }
                        else
                        {
                            distance = GetDistance(x1, y1, z1, x2, y2, z2);
                            runTime += GetTime(distance, moveSpeedXY);
                        }
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.StartsWith("J3"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //Console.WriteLine($"{subs[0]} {subs[1]} {subs[2]} {subs[3]}");
                        x2 = Convert.ToDouble(subs[1]);
                        y2 = Convert.ToDouble(subs[2]);
                        z2 = Convert.ToDouble(subs[3]);
                        if (IsZ(x1, y1, x2, y2))
                        {
                            distance = Math.Abs(z2 - z1);
                            runTime += GetTime(distance, moveSpeedZ);
                        }
                        else
                        {
                            distance = GetDistance(x1, y1, z1, x2, y2, z2);
                            runTime += GetTime(distance, moveSpeedXY);
                        }
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.StartsWith("CG"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        //double startX = 0, startY = 0, endX = 0, endY = 0, xOffset = 0, yOffset = 0;
                        double startX = x1;
                        double startY = y1;
                        double endX = Convert.ToDouble(subs[1]);
                        double endY = Convert.ToDouble(subs[2]);
                        double xOffset = Convert.ToDouble(subs[3]);
                        //CG variables
                        double yOffset = Convert.ToDouble(subs[4]);

                        distance = GetArcLength(startX, startY, endX, endY, xOffset, yOffset);

                        runTime += GetTime(distance, moveSpeedXY);

                        x1 = endX;
                        y1 = endY;
                    }
                    else if (line.StartsWith("JH"))
                    {
                        x1 = 0;
                        y1 = 0;
                        distance = GetDistance(x1, y1, z1, x2, y2, z2);
                        runTime += GetTime(distance, moveSpeedXY);
                        x1 = x2;
                        y1 = y2;
                        z1 = z2;
                    }
                    else if (line.Contains("SafeZ"))
                    {
                        string[] subs = line.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        z1 = Convert.ToDouble(subs[subs.Length - 1]);
                    }
                    else if (line.StartsWith("END"))
                    {
                        break;
                    }
                }

                //Add current run time to total
                totRunTime += runTime;

                //Add to the list of objects to return
                runTimeObjects.Add(new RunTime(filePaths[nameCount], (int)runTime));

                //Sets it so the next name received for the current one being processed is one over
                nameCount++;

                //Resets runtime to 0
                runTime = 0;
            }

            return runTimeObjects;
        }

        static double GetArcLength(double x1, double y1, double x2, double y2, double xOffset, double yOffset)
        {
            double r, d, c1, c2, theta, arcLength;
            c1 = x1 + xOffset;
            c2 = y1 + yOffset;

            r = Math.Sqrt(Math.Pow((c1 - x1), 2) + Math.Pow((c2 - y1), 2));
            d = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2));
            theta = Math.Acos(((2 * Math.Pow(r, 2)) - Math.Pow(d, 2)) / (2 * Math.Pow(r, 2)));
            arcLength = r * theta;

            return arcLength;
        }

        static double GetTime(double distance, double speed)
        {
            double time = (distance / speed);
            return Math.Abs(time);
        }

        static double GetDistance(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            double distance = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2) + Math.Pow((z2 - z1), 2));
            return Math.Abs(distance);
        }

        static bool IsZ(double x1, double y1, double x2, double y2)
        {
            double x = x2 - x1;
            double y = y2 - y1;
            if (x == 0 && y == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
