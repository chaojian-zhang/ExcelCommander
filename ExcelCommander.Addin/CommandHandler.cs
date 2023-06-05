using ExcelCommander.Base.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelCommander.Addin
{
    internal class CommandHandler
    {
        #region Entry Point
        internal CommandData Handle(CommandData data)
        {
            MessageBox.Show(data.Contents, data.CommandType);
            return null;
        }
        #endregion

        #region Handling Routines

        #endregion
    }
}
