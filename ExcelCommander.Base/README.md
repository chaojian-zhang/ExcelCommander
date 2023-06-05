# ExcelCommander.Base

This .Net Standard 2.0 class library defines the most basic types shared by the Excel Add-in and the client program. It's used as a bridge to connect both worlds.

It mostly contains type definitions and should not contain processing logic.

Avoid referencing any packages here so as to avoid potential runtime issues with loading assemblies in Excel.