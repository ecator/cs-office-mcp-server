using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
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
