using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuGetMsbuildChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            var version = args[0];
            MsbuildChecker.CheckVersion(version);
        }
    }
}
