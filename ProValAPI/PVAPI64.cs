using System;
using System.Runtime.InteropServices;

namespace ProValAPI
{
    /// <summary>
    /// PVAPI64.dll (ProVal's PVAPI)
    /// </summary>

    [ProgId("PVAPI64")]
    //[Guid("CBEDF2F4-E4F3-4693-8637-DCE680160643")] // this is the Guid for 32-bit dll
    [Guid("368C1BB8-6689-4AAE-9FFF-AE86F4C6EFBB")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PVAPI64
    {
        public APLW.WSEngine ws; // global workspace 

        [ComVisible(true)]

        // ------ Class properties -------- //
        //          read-only
        public string SysDir => ws.Call("SYSDIR");
        public string UserDir => ws.Variable["USERDIR"];
        public string ProVer => ws.Exec("PROVER");
        public string APLVer => ws.SysCall("SYSVER");
        public string CurrClient => ws.Variable["LIBDIR"];

        // -------- Class Constructor -------
        public PVAPI64()
        {
            try
            {
                ws = new APLW.WSEngine();

                LogEvent.Log("");
                LogEvent.Log("Created instance of APL WS Engine ...");
                LogEvent.Log("ProVal application directory: " + SysDir);
                LogEvent.Log("Working/User directory: " + UserDir);
                ws.Visible = 1; // for debug only set to true
            }
            catch (Exception e)
            {
                LogEvent.Log("Error: " + e.ToString());
            }
        }

        // --------- Class destructor -------
        ~PVAPI64()
        {
            try
            {
                LogEvent.Log("Closing the WSEngine.");
                ws.Close();
            }
            catch (Exception e)
            {
                LogEvent.Log("Error: " + e.ToString());
            }
        }

        // --------- Methods ------------- //
        public int Add(int x, int y) => ws.Exec(x + " + " + y);  // test APL function that adds two integers

        public dynamic PVCall(string sFuncname, dynamic xParamList = null)
        {
            try
            {
                if (xParamList == null)
                {
                    xParamList = "";
                }
                LogEvent.Log(sFuncname + " called");

                var RetVal = ws.Call("API_Call", xParamList, sFuncname);

                if (RetVal != null)
                {
                    for (int i = 0; i < RetVal.Length; i++)
                    {
                        var res = RetVal[i];
                        if (res is Array a)
                        {
                            if (a.Rank == 2)
                            {
                                Console.WriteLine($"res[{i}]={a}");
                                RetVal[i] = FormatArr(a);
                            }
                        }
                    }
                }
                else
                {
                    if (sFuncname.ToUpper() != "SHUTDOWN")
                    {
                        LogEvent.Log("WARNING: The function " + sFuncname + " returned no value.");
                    }
                }
                return RetVal;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                throw;
            }
        }

        private dynamic FormatArr(dynamic p)
        {
            Console.WriteLine("Formatting the array");

            if (p is int[,] intArr)
            {
                var nRows = intArr.GetUpperBound(0) + 1;
                var nCols = intArr.GetUpperBound(1) + 1;
                var res = new int[nCols, nRows];
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        res[j, i] = intArr[i, j];
                    }
                }
                return res;
            }

            if (p is double[,] doubleArr)
            {
                var nRows = doubleArr.GetUpperBound(0) + 1;
                var nCols = doubleArr.GetUpperBound(1) + 1;
                var res = new double[nCols, nRows];
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        res[j, i] = doubleArr[i, j];
                    }
                }
                return res;
            }
            if (p is object[,] objectArr)
            {
                var nRows = objectArr.GetUpperBound(0) + 1;
                var nCols = objectArr.GetUpperBound(1) + 1;
                var res = new object[nCols, nRows];
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        res[j, i] = objectArr[i, j];
                    }
                }
                return res;
            }

            if (p is bool[,] boolArr)
            {
                var nRows = boolArr.GetUpperBound(0) + 1;
                var nCols = boolArr.GetUpperBound(1) + 1;
                var res = new bool[nCols, nRows];
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        res[j, i] = boolArr[i, j];
                    }
                }
                return res;
            }
            return p;
        }
        // ----- consider transposing higher dimension arrays as phase 2
        // ----- for phase 1, we just want to replicate what was done
    }
}
