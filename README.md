# ExcelCommander

## Architecture

There are four distinct uses

* Repl interactively using either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Write text-based scripts and execute in either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in C# through Pure or Nugets.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in Python through PythonNet.