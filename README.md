## Recycle
This module adds Recycle Bin capabilities to PowerShell.

By default, PowerShell's Remove-Item function deletes files and objects completely, without a way to recover them. This module attempts to correct that by allowing files and directories to be deleted to the Windows Recycle Bin, instead.

_Note:_ This module does import the Visual Basic library in order to accomplish this.
## Usage
  Usage: Recycle [Path or Piped Object] <-Recurse> <-Force> <-WhatIf> <-Confirm> <-Help>

  • Accepts piped objects, such as those sent from Get-Item or Get-ChildItem, for example: 

      Get-ChildItem C:\Temp | Recycle      
      Get-Item C:\Temp\file1.txt, C:\TMP | Recycle
      Get-Item C:\Temp\file1.txt, C:\TMP, C:\SomeDirectoryTree | Recycle -Recurse

  • Since piped objects are allowed, error-checking is included to make sure the objects passed to this function are either files or directories.

  • As a safety measure, the -Recurse option must be used if the object passed to it is a directory.

  • Supports the -WhatIf switch to demonstrate what would happen for each action.

  • Suppports the -Confirm switch as extra protection for activities deemed to have a higher risk impact.

  • The -Force switch will exit the function with an exception if an interrupt is encountered. Otherwise, interrupts fail silently.

  • Also, when deleting Read-Only files the -Force switch is required.
