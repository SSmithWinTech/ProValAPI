using System;
using System.Runtime.InteropServices;

namespace ProValAPI
{
    /// <summary>
    /// PVAPI64.dll (ProVal's API (64-bit))
    /// </summary>

    [ProgId("PVAPI64")]
    [Guid("368C1BB8-6689-4AAE-9FFF-AE86F4C6EFBB")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PVAPI64
    {
        private APLW.WSEngine ws; // global workspace 

        [ComVisible(true)]

        // ------ Class properties -------- //
        //          read-only
        public string SysDir => ws.Call("SYSDIR");
        public string UserDir => ws.Variable["USERDIR"];
        public string ProVer => ws.Exec("PROVER");
        public string APLVer => ws.SysCall("SYSVER");
        public string CurrClient => ws.Variable["LIBDIR"];

        private readonly bool debug = false;

        // -------- Class Constructor -------
        public PVAPI64()
        {
            try
            {
                LogEvent.Log("before ws engine created");
                ws = new APLW.WSEngine();
                LogEvent.Log("after ws engine created");
                debug = DebugMode();
                if ( debug ) { 
                    LogEvent.Log("Created instance of APL WS Engine ...");
                    LogEvent.Log("ProVal application directory: " + SysDir);
                    LogEvent.Log("Working/User directory: " + UserDir);
                    LogEvent.Log("ProVal version: " + ProVer);
                    //if (VisibleMode())
                    //{
                    //    ws.Visible = 1;
                    //    // make sure the development version is being used in order for it to be visible
                    //}
                }
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
                if ( debug )
                { 
                    LogEvent.Log("Closing the WSEngine.");
                }
                ws.Close();
            }
            catch (Exception e)
            {
                if ( debug )
                { 
                    LogEvent.Log("Error: " + e.ToString());
                }
            }
        }

        // --------- Methods ------------- //
        public int Add(int x, int y) => ws.Exec(x + " + " + y);  // test APL function that adds two integers

        bool DebugMode()
        {
            var debug = ws.Exec("INIGET '[API] DEBUG'");
            if (debug is int && 1==(int)debug)
            {
                LogEvent.Log("");
                LogEvent.Log("Debug=1");
                return true;
            }
            return false;
            
        }

        bool VisibleMode()
        {
            var visible = ws.Exec("INIGET '[API] VISIBLE'");
            if (visible is int && 1 == (int)visible)
            {
                LogEvent.Log("Visible=1");
                return true;
            }
            return false;
        }

        public dynamic PVCall(string sFuncname, dynamic xParamList = null)
        {
            try
            {
                if (xParamList == null)
                {
                    xParamList = "";
                }
                if ( debug )
                {
                    LogEvent.Log(sFuncname + " called");
                }                

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
                                RetVal[i] = FormatArr(a);
                            }
                        }
                    }
                }
                else
                {
                    if (sFuncname.ToUpper() != "SHUTDOWN")
                    {
                        if ( debug ) { 
                            LogEvent.Log("WARNING: The function " + sFuncname + " returned no value.");
                        }
                    }
                }
                return RetVal;
            }
            catch (Exception e)
            {
                if (debug) { 
                    LogEvent.Log("ERROR:  " + e.ToString());
                }
                throw;
            }
        }

        private dynamic FormatArr(dynamic p)
        {
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
