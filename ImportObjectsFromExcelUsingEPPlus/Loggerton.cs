using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ImportObjectsFromExcelUsingEPPlus
{

    [Flags]
    public enum EnumLogFlags
    {
        None = 0,
        Information = 1,
        Event = 2,
        Warning = 4,
        Error = 8,
        All = 0xffff
    }

    /// <summary>
    /// A single entry into the log.
    /// </summary>
    public class LogEntry
    {
        public EnumLogFlags Flags { get; set; }

        public DateTime TimeStamp { get; set; }

        public string Message { get; set; }

        /// <summary>
        /// Has this message been excluded by regex excludes?
        /// </summary>
        public bool IsExcluded { get; set; }

        public LogEntry(EnumLogFlags flags, String msg)
        {
            Flags = flags;
            TimeStamp = DateTime.Now;
            Message = msg;
        }

        public override string ToString()
        {
            return $"{Flags}:{TimeStamp.ToString("HH:mm:ss.ff")} {Message}";
        }
    }

    /// <summary>
    /// The loggerton singleton.
    /// </summary>
    public sealed class Loggerton
    {
        private static readonly Loggerton instance = new Loggerton();
        public static Loggerton Instance { get { return instance; } }

        static Loggerton()
        {
        }
        private Loggerton()
        {

        }

        public bool IsEnabled = true;

        private List<LogEntry> Logs = new List<LogEntry>();

        /// <summary>
        /// An internal list of the regex expressions that indicate which
        /// logs to exclude. See 'SetExcludes'
        /// </summary>
        /// <seealso cref="SetExcludes(string)"/>
        private List<string> excludesList = new List<string>();

        /// <summary>
        /// A comma list of simple regex expressions to use to exclude logentries
        /// from the log.
        /// </summary>
        /// <param name="commalist"></param>
        public void SetExcludes(string commalist)
        {
            excludesList.Clear();
            excludesList = commalist.Split(',').ToList();

            // Reevaluate all the logs
            foreach (LogEntry le in Logs)
            {
                le.IsExcluded = false;
                foreach (string expr in excludesList)
                {
                    if (Regex.IsMatch(le.Message, expr, RegexOptions.IgnoreCase))
                    {
                        le.IsExcluded = true;
                        goto GetNextLogEntry;
                    }
                }
                GetNextLogEntry:;
            }
        }

        /// <summary>
        /// Get the filtered logs (non-excluded and orderded by timestamps)
        /// </summary>
        /// <param name="flags"></param>
        /// <param name="excludes"></param>
        /// <returns></returns>
        public string GetLogs(EnumLogFlags flags)
        {
            List<LogEntry> filteredLogs = Logs
                .Where(rr => (rr.Flags & flags) != 0)
                .Where(rr => !rr.IsExcluded)
                .OrderByDescending(rr => rr.TimeStamp)
                .ToList();

            StringBuilder sb = new StringBuilder();
            foreach (LogEntry le in filteredLogs)
            {
                sb.AppendLine($"{le}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Clear all log entries.
        /// </summary>
        public void ClearLogs()
        {
            Logs.Clear();
        }

        /// <summary>
        /// Write the logs to a file.
        /// </summary>
        /// <param name="path"></param>
        public void WriteLogs(string path)
        {
            try
            {
                File.WriteAllText(path, Logs.ToString());
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Cannot Write to path={path}. Err={ex.Message}");
            }

        }

        /// <summary>
        /// Log a message.
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="message"></param>
        public void LogIt(EnumLogFlags logType, string message)
        {
            if (!IsEnabled)
                return;

            bool isExcluded = false;
            foreach (string expr in excludesList)
            {
                if (Regex.IsMatch(message, expr, RegexOptions.IgnoreCase))
                {
                    isExcluded = true;
                    break;
                }
            }

            LogEntry entry = new LogEntry(logType, message);
            entry.IsExcluded = isExcluded;
            Logs.Add(entry);
        }

        /// <summary>
        /// Log a message with the default 'Information' type.
        /// </summary>
        /// <param name="message"></param>
        public void LogIt(string message)
        {
            LogIt(EnumLogFlags.Information, message);
        }

    }  // class
}





