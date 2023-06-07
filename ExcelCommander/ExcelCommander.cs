using ExcelCommander.Base.ClientServer;
using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;
using System.Drawing;

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

            Console.WriteLine($"Service started at port {Port}.");
        }
        #endregion

        #region Disposal
        public void Dispose()
        {
            Client.Close();
        }
        #endregion

        #region Handling
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
            Dispose();
        }
        public void ExecuteCommand(string command)
        {
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
                Contents = $"{nameof(SetFontSize)} {range} {color}"
            });
            return null;
        }
        public CommandData Bold(string range, string weight)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetValueFormat)} {range} {weight}"
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
        public CommandData Clear(string range)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Clear)} {range}"
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
        public CommandData SetValueFormat(string range, string format)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetValueFormat)} {range} \"{format}\""
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
        public CommandData SetFontColor(string range, string color)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetFontColor)} {range} {color}"
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
        public CommandData SetEquation(string cell, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetEquation)} {cell} \"{equation}\""
            });
            return null;
        }
        public CommandData SetEquation(int row, int col, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetEquation)} {row} {col} \"{equation}\""
            });
            return null;
        }
        public CommandData SetEquation(string row, string col, string equation)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(SetEquation)} {row} {col} \"{equation}\""
            });
            return null;
        }
        public CommandData Set(string range, double value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Set)} {range} {value}"
            });
            return null;
        }
        public CommandData Set(string range, int value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Set)} {range} {value}"
            });
            return null;
        }
        public CommandData Set(string range, string value)
        {
            Client.Send(new CommandData
            {
                CommandType = "Development",
                Contents = $"{nameof(Set)} {range} \"{value}\""
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
        #endregion

        #region Interface Calls - Management
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
    }
}
