# OfficeScrubC2R - Development Guidelines

## Code Quality Standards

### Naming Conventions
- **Classes**: PascalCase with descriptive names (e.g., `RegistryHelper`, `WindowsInstallerHelper`)
- **Methods**: PascalCase with action-oriented names (e.g., `DeleteKey`, `EnumerateValues`, `TerminateProcesses`)
- **Properties**: PascalCase for public properties (e.g., `Registry`, `Files`, `Processes`)
- **Constants**: UPPER_CASE with underscores (e.g., `OFFICE_ID`, `PROD_LEN`, `KEY_WOW64_64KEY`)
- **Fields**: camelCase with underscore prefix for private fields (e.g., `_regHelper`, `_pendingDeletes`, `_is64Bit`)
- **Parameters**: camelCase (e.g., `compressedGuid`, `directoryPath`, `processNames`)

### Code Organization Patterns
- **Region-based organization**: Use `#region` blocks to group related functionality
  ```csharp
  #region Enums and Constants
  #region P/Invoke Declarations  
  #region Registry Operations
  #region File Operations
  ```
- **Logical grouping**: Separate concerns into distinct helper classes (Registry, File, Process, Shell, etc.)
- **Static utility classes**: Use static classes for utility functions (e.g., `GuidHelper`, `OfficeConstants`)

### Error Handling Standards
- **Silent failure pattern**: Use try-catch blocks that return false/null on failure without throwing
  ```csharp
  try
  {
      // Operation
      return true;
  }
  catch
  {
      return false;
  }
  ```
- **Resource disposal**: Always use `using` statements or explicit disposal in finally blocks
- **Graceful degradation**: Provide fallback mechanisms when primary operations fail

## Architectural Patterns

### Helper Class Pattern
- Create specialized helper classes for different system areas (Registry, File, Process, etc.)
- Each helper class encapsulates related operations and maintains internal state
- Use dependency injection pattern for helper composition in orchestrator class

### P/Invoke Integration Pattern
- Group all Win32 API declarations in internal static `NativeMethods` class
- Use proper marshaling attributes and error handling for native calls
- Define constants for Win32 API flags and error codes within the same class

### WOW64 Support Pattern
- Always check both 32-bit and 64-bit registry views on 64-bit systems
- Use `_is64Bit` field to conditionally execute WOW64-specific operations
- Implement `GetWow64Key()` method to transform registry paths for 32-bit view

### Parallel Processing Pattern
- Use `Task.Run()` for CPU-intensive operations that can be parallelized
- Implement thread-safe collections with lock statements for shared data
- Use `Task.WaitAll()` with timeouts to prevent indefinite blocking

## Implementation Standards

### Registry Operations
- Always handle both regular and WOW64 registry views
- Use HashSet with StringComparer.OrdinalIgnoreCase for key/value collections
- Implement recursive key deletion with proper error handling
- Cache registry helper instance for performance

### File System Operations
- Remove read-only attributes before attempting deletion
- Use cmd.exe `rd /s /q` for performance-critical directory deletion
- Implement reboot scheduling for locked files using `MoveFileEx`
- Track pending deletes in internal collections

### Process Management
- Use parallel task execution for process termination
- Implement timeout mechanisms for process operations
- Always dispose Process objects in finally blocks
- Handle process access exceptions gracefully

### COM Integration
- Use dynamic typing for Shell COM objects to avoid version dependencies
- Support multiple localized verb names for international compatibility
- Implement proper COM object cleanup and resource management

## Performance Optimization Patterns

### StringBuilder Usage
- Use StringBuilder with pre-allocated capacity for string building operations
- Prefer StringBuilder over string concatenation in loops

### Lazy Evaluation
- Use LINQ methods like `Any()` for early termination in collection searches
- Implement short-circuit evaluation in boolean operations

### Resource Pooling
- Reuse helper instances across operations
- Cache frequently accessed registry keys and values

### Batch Operations
- Group related registry operations together
- Use parallel processing for independent operations

## Security and Safety Patterns

### Input Validation
- Always validate string lengths and formats before processing
- Check for null or empty inputs at method entry points
- Validate GUID formats before attempting conversion operations

### Privilege Handling
- Assume administrator privileges are available
- Implement graceful degradation when operations require elevated access
- Use appropriate registry access rights (KEY_READ, KEY_WRITE, KEY_ALL_ACCESS)

### Safe String Operations
- Use Path.GetFileNameWithoutExtension() for process name extraction
- Implement safe substring operations with bounds checking
- Use StringComparison.OrdinalIgnoreCase for case-insensitive comparisons

## Testing and Debugging Patterns

### Defensive Programming
- Return early on invalid inputs
- Use null-conditional operators where appropriate
- Implement bounds checking for array and string operations

### Logging Integration Points
- Design methods to return success/failure indicators
- Provide detailed operation results for external logging
- Track pending operations for post-execution reporting

### Testability Design
- Separate business logic from system calls
- Use dependency injection for testable components
- Design methods with single responsibilities

## Documentation Standards

### XML Documentation
- Document public methods with summary, parameters, and return values
- Include usage examples for complex operations
- Document known limitations and requirements

### Code Comments
- Use inline comments sparingly, prefer self-documenting code
- Comment complex algorithms and business logic
- Explain Win32 API usage and parameter meanings

### Constant Documentation
- Document the purpose and source of magic numbers
- Explain registry paths and their significance
- Document known GUID patterns and their meanings