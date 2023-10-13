using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal interface IReading
    {
        List<Vagt> ReadData(string target, string FilePath);
    }
}
