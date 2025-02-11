# SPS-WinTools Module

First I want to thanks [@guyrleech](https://github.com/guyrleech) for his script `Get-ProcessRelatives` which I have integrated in the SPS-WinTools module. I have made some changes to the script to make it more modular.

## Overview

The `SPS-WinTools` module provides a set of tools to retrieve and manage process information on Windows systems. One of the key functions in this module is to get parent and child process details recursively.

# Get Process Relatives

## Synopsis

Get parent and child processes details and recurse.

## Description

Uses `win32_process`. Level 0 processes are those specified via parameter, positive levels are parent processes & negative levels are child processes. Child processes are not guaranteed to be directly below their parent - check process id and parent process id.

## Parameters

- **Name**: A regular expression to match the name(s) of the process(es) to retrieve.
- **Id**: The ID(s) of the process(es) to retrieve.
- **IndentMultiplier**: The multiplier for the indentation level. Default is 1.
- **Indenter**: The character(s) used for indentation. Default is a space.
- **UnknownProcessName**: The placeholder name for unknown processes. Default is `<UNKNOWN>`.
- **Properties**: The properties to retrieve for each process.
- **Quiet**: Suppresses warning output if specified.
- **Norecurse**: Prevents recursion through processes if specified.
- **NoIndent**: Disables creating indented name if specified.
- **NoChildren**: Excludes child processes from the output if specified.
- **NoOwner**: Excludes the owner from the output if specified which speeds up the script.
- **SessionId**: -1 for same session as script, 0 for all sessions, or a positive integer for a specific session.

## Examples

### Example 1

```powershell
Get-ProcessRelatives -id 12345
```
Get parent and child processes of the running process with PID 12345.

### Example 2
```powershell
Get-ProcessRelatives -name notepad.exe,winword.exe -properties *
```
Get parent and child processes of all running processes of `notepad` and `winword`, outputting all `win32_process` properties & added ones.

### Example 3
```powershell
Get-ProcessRelatives -name powershell.exe -sessionid -1
```
Get parent and child processes of all running `powershell` processes in the same session as the script.

### Notes
* Modification history
    * 2024/09/13 @guyrleech Script born (Get Process Relatives)
    * 2024/09/16 @guyrleech First release (Get Process Relatives)
    * 2024/09/17 @SwissPowerShell integrated in SPS-WinTools module (Get Process Relatives)
    * 2024/02/11 @SwissPowerShell added Get-EnumInfo & Get-TypeInfo

### License
This module is licensed under the MIT License. See the LICENSE file for more details.

### Contributing
Contributions are welcome! Please submit a pull request or open an issue to discuss your changes.