using ExcelCommander.Base.ClientServer;
using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;

namespace ExcelCommander
{
    public sealed class ExcelCommander : IDisposable, ICommander
    {
        #region Construction
        private int Port { get; }
        private Client Client { get; set; }
        public static ExcelCommander Connect(int port)
            => new ExcelCommander(port);
        public ExcelCommander(int port)
        {
            Port = port;

            Client = new Client(Port, data => null);
            Client.Start();

            Console.WriteLine($"Commander connected to port {Port}.");
        }
        #endregion

        #region Disposal
        public void Dispose()
        {
            Client.Close();
        }
        #endregion

        #region Handling
        public void ExecuteFile(string scriptPath)
        {
            if (System.IO.File.Exists(scriptPath))
            {
                var lines = File.ReadAllLines(Path.GetFullPath(scriptPath));
                Execute(lines, false);
            }
        }
        public void Execute(string scripts)
        {
            Execute(scripts.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries), false);
        }
        public void Execute(string[] commands, bool interpretIfNull = true)
        {
            if (commands == null && interpretIfNull)
            {
                while (true)
                {
                    Console.Write("> ");
                    string input = Console.ReadLine();
                    ExecuteCommand(input);
                }
            }
            else
            {
                foreach (var command in commands)
                    ExecuteCommand(command);
            }
        }
        public void ExecuteCommand(string command)
        {
            if (string.IsNullOrWhiteSpace(command) || command.StartsWith("//") || command.StartsWith('#'))
                return;

            if (command == "Help")
                Console.WriteLine(CommanderHelper.GetHelpString());
            else if (command.StartsWith("Get"))
            {
                CommandData reply = Client.SendAndReceive(new CommandData
                {
                    CommandType = "Development",
                    Contents = command
                });
                Console.WriteLine($"[{reply.CommandType}] {reply.Contents}");
            }
            else
            {
                Client.Send(new CommandData
                {
                    CommandType = "Development",
                    Contents = command
                });
            }
        }
        #endregion

        #region Interface Calls - Reading
        public CommandData Get(string range)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Get)} {range}"
            });
        }
        public CommandData GetCell(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCell)} {cell}"
            });
        }
        public CommandData GetCell(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCell)} {row} {col}"
            });
        }
        public CommandData GetCell(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCell)} {row} {col}"
            });
        }
        public CommandData GetCellColor(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellColor)} {cell}"
            });
        }
        public CommandData GetCellColor(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellColor)} {row} {col}"
            });
        }
        public CommandData GetCellColor(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellColor)} {row} {col}"
            });
        }
        public CommandData GetCellName(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellName)} {cell}"
            });
        }
        public CommandData GetCellName(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellName)} {row} {col}"
            });
        }
        public CommandData GetCellName(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellName)} {row} {col}"
            });
        }
        public CommandData GetCellFontWeight(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFontWeight)} {cell}"
            });
        }
        public CommandData GetCellFontWeight(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFontWeight)} {row} {col}"
            });
        }
        public CommandData GetCellFontWeight(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFontWeight)} {row} {col}"
            });
        }
        public CommandData GetCellValueFormat(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValueFormat)} {cell}"
            });
        }
        public CommandData GetCellValueFormat(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValueFormat)} {row} {col}"
            });
        }
        public CommandData GetCellValueFormat(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValueFormat)} {row} {col}"
            });
        }
        public CommandData GetCellValue(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValue)} {cell}"
            });
        }
        public CommandData GetCellValue(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValue)} {row} {col}"
            });
        }
        public CommandData GetCellValue(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValue)} {row} {col}"
            });
        }
        public CommandData GetCellFormula(string cell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFormula)} {cell}"
            });
        }
        public CommandData GetCellFormula(int row, int col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFormula)} {row} {col}"
            });
        }
        public CommandData GetCellFormula(string row, string col)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellFormula)} {row} {col}"
            });
        }
        public CommandData GetCellValues(string cell, int rows, int cols)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValues)} {cell} {rows} {cols}"
            });
        }
        public CommandData GetCellValues(string cell, string rows, string cols)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValues)} {cell} {rows} {cols}"
            });
        }
        public CommandData GetCellValues(string startcell, string endcell)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValues)} {startcell} {endcell}"
            });
        }
        public CommandData GetCellValues(string range)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCellValues)} {range}"
            });
        }
        public CommandData GetTable(string tableName)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetTable)} \"{tableName}\""
            });
        }
        public CommandData GetSheet(string sheetName)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetSheet)} \"{sheetName}\""
            });
        }
        public CommandData GetCurrentSheet()
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetCurrentSheet)}"
            });
        }
        public CommandData GetSheets()
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GetSheets)}"
            });
        }
        public CommandData HasSheet(string name)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(HasSheet)} \"{name}\""
            });
        }
        public CommandData HasTable(string name)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(HasTable)} \"{name}\""
            });
        }
        public CommandData HasNamedRange(string name)
        {
            return Client.SendAndReceive(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(HasNamedRange)} \"{name}\""
            });
        }
        #endregion

        #region Interface Calls - Writing
        public CommandData Align(string range, string option)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Align)} {range} {option}"
            });
            return null;
        }
        public CommandData Border(string range, string weight)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetFontSize)} {range} {weight}"
            });
            return null;
        }
        public CommandData Background(string range, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Background)} {range} {color}"
            });
            return null;
        }
        public CommandData Bold(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Bold)} {range}"
            });
            return null;
        }
        public CommandData Bold(string range, string toggle)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Bold)} {range} {toggle}"
            });
            return null;
        }
        public CommandData Cell(string range, double value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Cell)} {range} {value}"
            });
            return null;
        }
        public CommandData Cell(string range, int value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Cell)} {range} {value}"
            });
            return null;
        }
        public CommandData Cell(string range, string value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Cell)} {range} \"{value}\""
            });
            return null;
        }
        public CommandData Clear(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Clear)} {range}"
            });
            return null;
        }
        public CommandData ClearAll()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(ClearAll)}"
            });
            return null;
        }
        public CommandData ClearFormat(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(ClearFormat)} {range}"
            });
            return null;
        }
        public CommandData CSV(string start, string filename)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(CSV)} {start} \"{filename}\""
            });
            return null;
        }
        public CommandData CreateSheet(string sheetName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(CreateSheet)} \"{sheetName}\""
            });
            return null;
        }
        public CommandData CreateTable(string range, string tableName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(CreateTable)} {range} \"{tableName}\""
            });
            return null;
        }
        public CommandData Fit(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Fit)} {range}"
            });
            return null;
        }
        public CommandData FitAll()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(FitAll)}"
            });
            return null;
        }
        public CommandData Italic(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Italic)} {range}"
            });
            return null;
        }

        public CommandData Italic(string range, string toggle)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Italic)} {toggle}"
            });
            return null;
        }
        public CommandData Merge(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Merge)} {range}"
            });
            return null;
        }
        public CommandData MoveSheetBefore(string sheetName, string otherSheetName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(MoveSheetBefore)} \"{sheetName}\" \"{otherSheetName}\""
            });
            return null;
        }
        public CommandData NameRange(string range, string rangeName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(NameRange)} {range} \"{rangeName}\""
            });
            return null;
        }
        public CommandData Outline(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Outline)} {range}"
            });
            return null;
        }
        public CommandData Color(string range, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Color)} {range} {color}"
            });
            return null;
        }
        public CommandData Color(int row, int col, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Color)} {row} {col} {color}"
            });
            return null;
        }
        public CommandData Color(string row, string col, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Color)} {row} {col} {color}"
            });
            return null;
        }
        public CommandData Formula(string cell, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Formula)} {cell} \"{equation}\""
            });
            return null;
        }
        public CommandData SetEquation(int row, int col, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Formula)} {row} {col} \"{equation}\""
            });
            return null;
        }
        public CommandData Formula(string row, string col, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Formula)} {row} {col} \"{equation}\""
            });
            return null;
        }
        public CommandData RenameSheet(string newName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(RenameSheet)} \"{newName}\""
            });
            return null;
        }
        public CommandData RenameSheet(string originalName, string newName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(RenameSheet)} \"{originalName}\" \"{newName}\""
            });
            return null;
        }
        public CommandData SetCell(string cell, int value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {cell} {value}"
            });
            return null;
        }
        public CommandData SetCell(string cell, double value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {cell} {value}"
            });
            return null;
        }
        public CommandData SetCell(string cell, string value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {cell} \"{value}\""
            });
            return null;
        }
        public CommandData SetCell(int row, int col, double value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {row} {col} {value}"
            });
            return null;
        }
        public CommandData SetCell(int row, int col, int value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {row} {col} {value}"
            });
            return null;
        }
        public CommandData SetCell(int row, int col, string value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {row} {col} \"{value}\""
            });
            return null;
        }
        public CommandData SetCell(string row, string col, string value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCell)} {row} {col} \"{value}\""
            });
            return null;
        }
        public CommandData SetCellName(string cell, string name)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCellName)} {cell} \"{name}\""
            });
            return null;
        }
        public CommandData SetCellName(string row, string col, string name)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetCellName)} {row} {col} \"{name}\""
            });
            return null;
        }
        public CommandData SetCellValues(string start, string csv)
        {
            Client.Send(new CommandData
            {
                CommandType = $"SetCellValues {start}",
                Contents = csv
            });
            return null;
        }
        public CommandData SetFontColor(string range, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetFontColor)} {range} {color}"
            });
            return null;
        }
        public CommandData SetFontSize(string range, int size)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetFontSize)} {range} {size}"
            });
            return null;
        }
        public CommandData SetFontSize(string range, string size)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetFontSize)} {range} {size}"
            });
            return null;
        }
        public CommandData SetValueFormat(string range, string format)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetValueFormat)} {range} \"{format}\""
            });
            return null;
        }
        public CommandData Width(string range, string width)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Width)} {range} {width}"
            });
            return null;
        }
        public CommandData Wrap(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Wrap)} {range}"
            });
            return null;
        }
        public CommandData Wrap(string range, string toggle)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Wrap)} {toggle}"
            });
            return null;
        }
        #endregion

        #region Interface Calls - Management
        public CommandData Select(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Select)} {range}"
            });
            return null;
        }
        public CommandData GoToSheet(string sheetName)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(GoToSheet)} \"{sheetName}\""
            });
            return null;
        }
        #endregion

        #region Macros
        public CommandData Apply()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Apply)}"
            });
            return null;
        }
        public CommandData Apply(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Apply)} {range}"
            });
            return null;
        }
        public CommandData Copy()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Copy)}"
            });
            return null;
        }
        public CommandData Duplicate()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Duplicate)}"
            });
            return null;
        }
        public CommandData Fill()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Fill)}"
            });
            return null;
        }

        public CommandData Fill(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Fill)} {range}"
            });
            return null;
        }
        public CommandData Fill(string range, string direction)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Fill)} {range} {direction}"
            });
            return null;
        }
        public CommandData Paste()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Paste)}"
            });
            return null;
        }
        public CommandData Paste(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Paste)} {range}"
            });
            return null;
        }
        public CommandData Save()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Save)}"
            });
            return null;
        }
        public CommandData Save(string outputFilePath)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Save)} \"{outputFilePath}\""
            });
            return null;
        }
        public CommandData Sort()
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Sort)}"
            });
            return null;
        }
        public CommandData Sort(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Sort)} {range}"
            });
            return null;
        }
        #endregion

        #region Programming
        public CommandData Evaluate(string scriptPath)
        {
            Execute(scriptPath);
            return null;
        }
        #endregion

        #region Utilities
        public CommandData Random(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Random)} {range}"
            });
            return null;
        }

        public CommandData Random(string range, string multiplier)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Random)} {range} {multiplier}"
            });
            return null;
        }

        public CommandData Random(string range, string from, string to)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Random)} {range} {from} {to}"
            });
            return null;
        }
        #endregion

        #region Finance
        public CommandData ETL(string range, string outputCell)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(ETL)} {range} {outputCell}"
            });
            return null;
        }
        public CommandData ETL(string range, string outputCell, string percentage)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(ETL)} {range} {outputCell} {percentage}"
            });
            return null;
        }
        #endregion
    }
}
