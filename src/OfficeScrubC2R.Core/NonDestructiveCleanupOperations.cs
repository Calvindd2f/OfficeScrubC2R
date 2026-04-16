using Microsoft.Win32;

namespace OfficeScrubC2R
{
    public static class NonDestructiveCleanupOperations
    {
        public static OperationResult PlanRegistryKeyDelete(RegistryHive hive, RegistryView view, string subKey)
        {
            return OperationResult.WouldRun(
                "Registry",
                "DeleteKey",
                "RegistryKey",
                subKey,
                "Registry deletion is represented as a dry-run operation in this baseline.",
                hive,
                view);
        }

        public static OperationResult PlanFileDelete(string path)
        {
            return OperationResult.WouldRun(
                "Files",
                "DeletePath",
                "FileSystemPath",
                path,
                "File deletion is represented as a dry-run operation in this baseline.");
        }

        public static OperationResult BlockDestructiveOperation(string step, string action, string targetKind, string target)
        {
            return OperationResult.Blocked(
                step,
                action,
                targetKind,
                target,
                "Destructive cleanup helpers are blocked until the full scrub implementation is added.",
                ScrubPlanner.DestructiveExecutionNotSupportedErrorId);
        }
    }
}
