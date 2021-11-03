using Antlr4.Runtime;

namespace FormExcel
{
    class Calculator
    {
        public static double Evaluate(string expression)
        {
            var lexer = new FormExcelLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());

            var tokens = new CommonTokenStream(lexer);
            var parser = new FormExcelParser(tokens);

            var tree = parser.compileUnit();

            var visitor = new LabCalculatorVisitor();

            return visitor.Visit(tree);
        }
    }
}

