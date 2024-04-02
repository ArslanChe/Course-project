using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Graphviz4Net.Dot.AntlrParser;
using Graphviz4Net.Dot;

namespace GraphLibrary
{
    public class GraphParser
    {
        public GraphParser() { }

        public DotGraph<string> ParseGraphData(string code)
        {
            var parser = AntlrParserAdapter<string>.GetParser();
            return parser.Parse(code);
        }
    }
}
