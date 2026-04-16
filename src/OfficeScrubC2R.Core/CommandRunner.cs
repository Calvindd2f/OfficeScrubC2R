using System;
using System.Diagnostics;

namespace OfficeScrubC2R
{
    public interface ICommandRunner
    {
        CommandRunResult Run(string fileName, string arguments, int timeoutMilliseconds);
    }

    public sealed class CommandRunResult
    {
        public CommandRunResult(int exitCode, string output)
        {
            ExitCode = exitCode;
            Output = output ?? string.Empty;
        }

        public int ExitCode { get; }
        public string Output { get; }
    }

    public sealed class ProcessCommandRunner : ICommandRunner
    {
        public CommandRunResult Run(string fileName, string arguments, int timeoutMilliseconds)
        {
            try
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = arguments,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    process.Start();
                    var outputTask = process.StandardOutput.ReadToEndAsync();
                    var errorTask = process.StandardError.ReadToEndAsync();

                    if (!process.WaitForExit(timeoutMilliseconds))
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch (Exception exception)
                        {
                            return new CommandRunResult(-1, fileName + " timed out. Process kill also failed: " + exception.GetType().Name + ": " + exception.Message);
                        }

                        return new CommandRunResult(-1, fileName + " timed out.");
                    }

                    var output = (outputTask.Result + Environment.NewLine + errorTask.Result).Trim();
                    return new CommandRunResult(process.ExitCode, output);
                }
            }
            catch (Exception exception)
            {
                return new CommandRunResult(-1, exception.GetType().Name + ": " + exception.Message);
            }
        }
    }
}
