using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelMAUIApp
{
    public static class Calculator
    {
        public static double Evaluate(string expression)
        {
                var lexer = new MyExcelMAUIAppLexer(new AntlrInputStream(expression));
                lexer.RemoveErrorListeners();
                lexer.AddErrorListener(new ThrowExceptionErrorListener());

                var tokens = new CommonTokenStream(lexer);
                var parser = new MyExcelMAUIAppParser(tokens);

                var tree = parser.compileUnit();

                var visitor = new MyExcelMAUIAppVisitor();

                return visitor.Visit(tree);
        }
    }
}
