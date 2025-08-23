using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;



namespace OfficeServer.Tools;

/// <summary>
/// Manages an COM Application instance and its associated COM objects, ensuring proper release.
/// Implements IDisposable for use with 'using' statements.
/// </summary>
public abstract class Session<TApplication> : IDisposable where TApplication : class
{

    protected TApplication Application { get; set; }
    protected List<object> _comObjectsToRelease = new List<object>();
    protected bool _disposed = false;


    /// <summary>
    /// Registers a COM object to be released when the session is disposed.
    /// </summary>
    /// <typeparam name="T">The type of the COM object.</typeparam>
    /// <param name="obj">The COM object instance.</param>
    /// <returns>The registered COM object.</returns>
    public T RegisterComObject<T>(T obj) where T : class
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            _comObjectsToRelease.Add(obj);
        }
        return obj;
    }

    /// <summary>
    /// Releases a single COM object.
    /// </summary>
    /// <param name="obj">The COM object to release.</param>
    private void ReleaseSingleComObject(object obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            try
            {
                // Loop to ensure all references are released
                while (Marshal.ReleaseComObject(obj) > 0) { }
            }
            catch (Exception ex)
            {
                // Log or handle the exception if needed, but don't rethrow
                Console.Error.WriteLine($"Warning: Failed to release COM object: {ex.Message}");
            }
        }
    }


    /// <summary>
    /// Run a macro.
    /// </summary>
    /// <param name="macroName">The name of macro</param>
    /// <param name="macroParameters">The parameters of macro. The maximum number is 30.</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public string RunMacro(string macroName, string[]? macroParameters = null)
    {
        dynamic response;

        dynamic app = Application;
        var macroParameterCnt = 0;
        if (macroParameters != null && macroParameters.Length > 0)
        {
            macroParameterCnt = macroParameters.Length;
        }
        try
        {
            switch (macroParameterCnt)
            {
                case 0:
                    response = app.Run(macroName);
                    break;
                case 1:
                    response = app.Run(macroName, macroParameters[0]);
                    break;
                case 2:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1]);
                    break;
                case 3:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2]);
                    break;
                case 4:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3]);
                    break;
                case 5:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4]);
                    break;
                case 6:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5]);
                    break;
                case 7:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6]);
                    break;
                case 8:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7]);
                    break;
                case 9:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8]);
                    break;
                case 10:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9]);
                    break;
                case 11:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10]);
                    break;
                case 12:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11]);
                    break;
                case 13:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12]);
                    break;
                case 14:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13]);
                    break;
                case 15:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14]);
                    break;
                case 16:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15]);
                    break;
                case 17:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16]);
                    break;
                case 18:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17]);
                    break;
                case 19:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18]);
                    break;
                case 20:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19]);
                    break;
                case 21:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20]);
                    break;
                case 22:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21]);
                    break;
                case 23:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22]);
                    break;
                case 24:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23]);
                    break;
                case 25:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24]);
                    break;
                case 26:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25]);
                    break;
                case 27:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26]);
                    break;
                case 28:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27]);
                    break;
                case 29:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27], macroParameters[28]);
                    break;
                case 30:
                    response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27], macroParameters[28], macroParameters[29]);
                    break;
                default:
                    throw new McpException("Exceeds the maximum number of macro parameters, which is 30.");
            }
        }
        catch (Exception ex)
        {
            throw new McpException(ex.Message);
        }
        if (string.IsNullOrEmpty(response))
        {
            return string.Empty;

        }
        return response.ToString();
    }


    /// <summary>
    /// Escape the value of the markdown table
    /// </summary>
    /// <param name="val">Value that requires escaping.</param>
    /// <returns></returns>
    public string EscapeMarkdownTableValue(string val)
    {
        if (string.IsNullOrEmpty(val))
        {
            return val;
        }
        val = val.Replace("\n", "<br>");
        val = val.Replace("\r", "");
        val = val.Replace("\\", "\\\\");
        val = val.Replace("|", "\\|");
        return val;
    }

    /// <summary>
    /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this); // Prevent finalizer from running
    }

    /// <summary>
    /// Releases all COM objects.
    /// </summary>
    /// <param name="disposing">True if called from Dispose(), false if called from finalizer.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            // Release managed resources here if any
        }

        if (_disposed) return;

        if (disposing)
        {
            // Release managed resources here if any
        }

        // Release unmanaged resources (COM objects)
        if (Application != null)
        {
            try
            {
                // Before quitting, set DisplayAlerts to false to avoid "Save changes?" prompts
                dynamic dynamicApp = Application;
                if (dynamicApp is Excel.Application)
                {
                    dynamicApp.DisplayAlerts = false;
                }
                else if (dynamicApp is Word.Application)
                {
                    dynamicApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                else if (dynamicApp is PowerPoint.Application)
                {
                    dynamicApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                dynamicApp.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: Failed to quit COM Application: {ex.Message}");
            }
            finally
            {
                // Release unmanaged resources (COM objects)
                for (int i = _comObjectsToRelease.Count - 1; i >= 0; i--)
                {
                    ReleaseSingleComObject(_comObjectsToRelease[i]);
                }
                _comObjectsToRelease.Clear();
            }
        }

        // Due to performance issues, commenting out explicit GC calls.
        // Explicitly call GC to clean up any remaining Runtime Callable Wrappers (RCWs)
        //GC.Collect();
        //GC.WaitForPendingFinalizers();

        _disposed = true;
    }

    /// <summary>
    /// Finalizer (destructor) in case Dispose is not called explicitly.
    /// </summary>
    ~Session()
    {
        Dispose(false);
    }
}
