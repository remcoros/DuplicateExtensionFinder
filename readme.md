# Duplicate Extension Finder for VS2015

Simple console tool to find duplicate extension folders / manifests for Visual Studio 2015 extensions. Used by me to cleanup and fix duplicate extension loading errors.

Extensions that mysteriously disable themselves can be an indicator of duplicate extensions.

Searches for extensions in

 * `%APP_DATA%\Microsoft\VisualStudio\14.0\Extensions`
 * `C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\Extensions`

## Command line parameters

 * _none_: shows all extensions installed (marking each line as "KEEP" or "DELETE")
 * `-dupes`: only show duplicate extensions
 * `-delete`: deletes the extension directory from the file system

## How to run

1. Clone/download and build
2. Run the .exe via a command prompt (CMD, Powershell) e.g.
      
         > DuplicateExtensionFinder.exe -dupes

    or

         > DuplicateExtensionFinder.exe -delete
  
    or run in Visual Studio by editing the "Command line arguments" found in _Project Properties => Debug => Start Options => Command line arguments_ then press `F5`
