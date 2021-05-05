using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using EnvDTE;
using EnvDTE80;
using EnvDTE90;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.Diagnostics;

namespace LLVM.ClangFormat
{
    public class ClangFormat
    {
        int m_dTopLine;
        int m_dBottomLine;
        int m_dCursorLine;
        int m_dRestoreCaret;
        string m_fullFileName;
        string m_sClangExecutable;
        string m_sCfStyle;
        string m_sFallbackStyle;
        string m_sAssumeFilename;
        string m_sCursorPosition;
        bool m_sSortIncludes;
        bool m_sSearchClangFormat;
        bool m_sSaveOnFormat;
        bool m_sOutputEnabled;
        bool m_formatLinesOnly;
        bool m_CfSuccessful = false;
        string cfOutput;
        string cfErrOutput;
        string m_dCurrentFileBuffer;
        EnvDTE.Properties mProps;
        String m_argument;
        StringBuilder tmpOut = new StringBuilder();
        StringBuilder cfErrOutputBuilder = new StringBuilder();
        ProcessStartInfo m_procStart = new ProcessStartInfo();
        OutputWindowPane m_owp;
        TextDocument m_td;

        public ClangFormat(OutputWindowPane owp, EnvDTE.Properties props)
        {
            mProps = props;

            string[] dirToSearch = { @"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\LLVM\LLVM", @"HKEY_LOCAL_MACHINE\SOFTWARE\LLVM" };

            m_sSearchClangFormat = (bool)mProps.Item("SearchClangFormat").Value;

            if (m_sSearchClangFormat)
            {
                try
                {
                    foreach(string registryDir in dirToSearch)
                    {
                        string llvmInstallationDir = (string)Registry.GetValue(registryDir, "", null);
                        if (llvmInstallationDir != null)
                        {
                            m_sClangExecutable = llvmInstallationDir + "\\bin\\clang-format.exe";
                            mProps.Item("ClangExecutable").Value = m_sClangExecutable;
                            break;
                        }
                    }
                }
                catch
                {
                    LogToOutputWindow("Caught exception, while trying to read registry and search for clang-format.");
                    LogToOutputWindow(Environment.NewLine);
                }
            }

            m_dTopLine = 0;
            m_dBottomLine = 0;
            m_dCursorLine = 0;
            m_dRestoreCaret = 0;
            m_fullFileName = "";
            cfOutput = "";
            cfErrOutput = "";
            m_dCurrentFileBuffer = "";
            m_sCfStyle = "";
            m_sFallbackStyle = "";
            m_sAssumeFilename = "";
            m_sCursorPosition = "Restore";
            m_sSortIncludes = true;
            m_sSearchClangFormat = true;
            m_sSaveOnFormat = true;
            m_sOutputEnabled = true;
            m_formatLinesOnly = false;
            m_CfSuccessful = false;
            m_argument = "";

            m_owp = owp;

            m_procStart.RedirectStandardInput = true;
            m_procStart.RedirectStandardOutput = true;
            m_procStart.RedirectStandardError = true;
            m_procStart.UseShellExecute = false;
            m_procStart.CreateNoWindow = true;
        }

        public bool FormatFile(EnvDTE.DTE dte)
        {
            m_CfSuccessful = false;
            DateTime localDate = DateTime.Now;
            string localDateString;
            string localTimeString;
            localDateString = String.Format("Date: {0:dddd, d. MMMM yyyy}", localDate);
            localTimeString = String.Format("Time: {0:HH:mm:ss}", localDate);
            //DateTime utcDate = DateTime.UtcNow;

            m_sClangExecutable = (string)mProps.Item("ClangExecutable").Value;
            m_sCfStyle = (string)mProps.Item("CfStyle").Value;
            m_sFallbackStyle = (string)mProps.Item("FallbackStyle").Value;
            m_sAssumeFilename = (string)mProps.Item("AssumeFilename").Value;
            m_sCursorPosition = (string)mProps.Item("CursorPosition").Value; ;
            m_sSortIncludes = (bool)mProps.Item("SortIncludes").Value;
            m_sSaveOnFormat = (bool)mProps.Item("SaveOnFormat").Value;
            m_sOutputEnabled = (bool)mProps.Item("OutputEnabled").Value;
            m_formatLinesOnly = false;
            m_argument = "";

            LogToOutputWindow(Environment.NewLine);
            LogToOutputWindow("[ Log " + localDateString + " ]");
            LogToOutputWindow(Environment.NewLine);
            LogToOutputWindow("[ Log " + localTimeString + " ]");
            LogToOutputWindow(Environment.NewLine);

            if (dte.ActiveDocument == null)
            {
                LogToOutputWindow("Make sure you have a document opened. Done nothing.");
                LogToOutputWindow(Environment.NewLine);
                return false;
            }

            if (dte.ActiveDocument.Type != "Text")
            {
                LogToOutputWindow("Document type is not a text document. Done nothing.");
                LogToOutputWindow(Environment.NewLine);
                return false;
            }

            Document tdToSave = dte.ActiveDocument;
            m_fullFileName = dte.ActiveDocument.FullName;  // full file name
            m_td = (TextDocument)dte.ActiveDocument.Object("");
            TextSelection sel = m_td.Selection;

            VirtualPoint topPt = sel.TopPoint;
            m_dTopLine = topPt.Line;
            VirtualPoint bottomPt = sel.BottomPoint;
            m_dBottomLine = bottomPt.Line;
            VirtualPoint activePt = sel.ActivePoint; // restore this point;
            m_dCursorLine = activePt.Line;
            int restoredCursorLine = m_dTopLine;
            m_dRestoreCaret = sel.ActivePoint.AbsoluteCharOffset;

            m_dCurrentFileBuffer = "";
            tmpOut.Length = 0;
            cfErrOutputBuilder.Length = 0;

            if (!(sel.IsEmpty))
            {
                m_formatLinesOnly = true;
            }
            // no selection
            sel.EndOfDocument(false);
            sel.StartOfDocument(true);
            //sel.SelectAll();
            m_dCurrentFileBuffer = sel.Text; // load complete buffer

            createArgumentString();
            m_procStart.Arguments = m_argument;

            LogToOutputWindow("Buffer of file: " + m_fullFileName);
            LogToOutputWindow(Environment.NewLine);

            startProcessAndGetOutput();

            try
            {
                //// write stdout from clang-format to editor buffer
                writeOutputToEditorBuffer(sel, tdToSave);
            }
            catch
            {
                LogToOutputWindow("Caught exception, while trying to change buffer.");
                LogToOutputWindow(Environment.NewLine);
            }

            // restore cursor
            if (m_sCursorPosition == "Top")
            {
                restoredCursorLine = 1;
                sel.MoveToLineAndOffset(restoredCursorLine, 1, false);
            }
            else if (m_sCursorPosition == "Bottom")
            {
                restoredCursorLine = m_td.EndPoint.Line;
                sel.MoveToLineAndOffset(restoredCursorLine, 1, false);
            }
            else if (m_sCursorPosition == "SameLine")
            {
                restoredCursorLine = m_dCursorLine;
                if (m_td.EndPoint.Line < restoredCursorLine)
                {
                    restoredCursorLine = m_td.EndPoint.Line;
                }
                sel.MoveToLineAndOffset(restoredCursorLine, 1, false);
            }
            else if (m_sCursorPosition == "Restore")
            {
                try
                {
                    sel.MoveToAbsoluteOffset(m_dRestoreCaret, false);
                }
                catch
                {}
            }

            return true;
        }

        private void createArgumentString()
        {
            if (File.Exists(m_sClangExecutable))
            {
                LogToOutputWindow("Executable: " + m_sClangExecutable);
                LogToOutputWindow(Environment.NewLine);
            }
            else
            {
                LogToOutputWindow("Make sure the path to the clang-format executable is correct");
                LogToOutputWindow(Environment.NewLine);
                LogToOutputWindow("Searched Executable at: " + m_sClangExecutable);
                LogToOutputWindow(Environment.NewLine);
            }

            m_argument = "";
            if (m_sCfStyle != "")
            {
                m_argument += "-style=" + m_sCfStyle;
            }
            else
            {
                m_argument += "-style=" + "file";
            }

            if (m_sFallbackStyle != "")
            {
                m_argument += " -fallback-style=" + m_sFallbackStyle;
            }
            else
            {
                m_argument += " -fallback-style=" + "none";
            }

            if (m_sAssumeFilename != "")
            {
                m_argument += " -assume-filename=" + m_sAssumeFilename;
            }
            else
            {
                m_argument += " -assume-filename=" + Path.GetFileName(m_fullFileName);
            }
            
            if (m_sSortIncludes)
            {
                m_argument += " -sort-includes";
            }
            else
            {
                m_argument += "";
            }

            if (m_formatLinesOnly)
            {
                m_argument += " -lines=" + m_dTopLine + ":" + m_dBottomLine;
            }
            else
            {
                m_argument += "";
            }

            if (m_sCursorPosition == "Restore")
            {
                m_argument += " -cursor=" + m_dRestoreCaret;
            }
            else
            {
                m_argument += "";
            }

            m_procStart.FileName = m_sClangExecutable;
            m_procStart.Arguments = m_argument;

            LogToOutputWindow("Arguments: " + m_argument);
            LogToOutputWindow(Environment.NewLine);
        }

        private void startProcessAndGetOutput()
        {
            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (System.Diagnostics.Process exeProcess = System.Diagnostics.Process.Start(m_procStart))
                {
                    exeProcess.StandardInput.Write(m_dCurrentFileBuffer);
                    exeProcess.StandardInput.Close();
                    while (!exeProcess.StandardOutput.EndOfStream)
                    {
                        tmpOut.Append(exeProcess.StandardOutput.ReadToEnd());
                    }
                    while (!exeProcess.StandardError.EndOfStream)
                    {
                        cfErrOutputBuilder.Append(exeProcess.StandardError.ReadToEnd());
                    }
                    cfOutput = tmpOut.ToString();
                    cfErrOutput = cfErrOutputBuilder.ToString();
                    if (m_sCursorPosition == "Restore" && (cfErrOutput.Length == 0))
                    {
                        setNewCursorPosition(ref cfOutput);
                        m_CfSuccessful = true;
                    }
                    else
                    {
                        LogToOutputWindow(Environment.NewLine);
                        LogToOutputWindow(cfErrOutput);
                        LogToOutputWindow(Environment.NewLine);
                        m_CfSuccessful = false;
                    }
                }
            }
            catch
            {
                LogToOutputWindow("Caught exception, while trying to execute clang-format or IO operations.");
                LogToOutputWindow(Environment.NewLine);
                m_CfSuccessful = false;
            }
        }

        private void writeOutputToEditorBuffer(TextSelection sel, Document tdToSave)
        {
            if (m_td.Type == "Text")
            {
                if (m_CfSuccessful)
                {
                    sel.Insert(cfOutput, (int)vsInsertFlags.vsInsertFlagsCollapseToEnd);
                    if (m_sSaveOnFormat)
                    {
                        tdToSave.Save(m_fullFileName);
                    }
                }
            }
        }

        private void LogToOutputWindow(string text)
        {
            if (m_sOutputEnabled)
            {
                m_owp.OutputString(text);
            }
        }

        private void setNewCursorPosition(ref string newBuffer)
        {
            string firstline = "";
            string[] array;
            try
            {
                int newLineIndex = newBuffer.IndexOf("\n");
                firstline = newBuffer.Substring(0, newLineIndex);
                newBuffer = newBuffer.Substring(newLineIndex + 1);
                array = firstline.Split(':');
                array = array[1].Split(',');
                string newLineNumber = array[0].Substring(1);
                m_dRestoreCaret = Int32.Parse(newLineNumber);
            }
            catch
            {
                LogToOutputWindow("Caught exception, while trying to parse clang-format output for new-line.");
                LogToOutputWindow(Environment.NewLine);
            }
        }
    }
}
