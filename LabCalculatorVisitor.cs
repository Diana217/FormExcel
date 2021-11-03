using Antlr4.Runtime.Misc;
using System.Collections.Generic;
using System.Diagnostics;

namespace FormExcel
{
    class LabCalculatorVisitor : FormExcelBaseVisitor<double>
    {
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();
        /*private IEnumerable<string> Separate(string input)
        {
            throw new NotImplementedException();
        }*/
        public override double VisitCompileUnit(FormExcelParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(FormExcelParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }

        //IdentifierExpr
        public override double VisitIdentifierExpr(FormExcelParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            //видобути значення змінної з таблиці
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }

        public override double VisitParenthesizedExpr(FormExcelParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitAdditiveExpr(FormExcelParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == FormExcelLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(FormExcelParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == FormExcelLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }
        public override double VisitMoDivExpr([NotNull] FormExcelParser.MoDivExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == FormExcelLexer.MOD)
            {
                return left % right;
            }
            else
            {
                return (int)left / (int)right;
            }
            //return base.VisitMoDivExpr(context);
        }
        public override double VisitUnaryPExpr(FormExcelParser.UnaryPExprContext context)
        {
            var number = WalkLeft(context);
            return number + 1;
        }

        public override double VisitUnaryMExpr(FormExcelParser.UnaryMExprContext context)
        {
            var number = WalkLeft(context);
            return number - 1;
        }
        public override double VisitMaxMinExpr(FormExcelParser.MaxMinExprContext context)
        {
            double x = WalkLeft(context);
            double y = WalkRight(context);
            if (context.operatorToken.Type == FormExcelParser.MAX)
            {
                Debug.WriteLine("MAX(");
                if (x > y)
                {
                    Debug.WriteLine(x);
                    return x;
                }
                else
                {
                    Debug.WriteLine(y);
                    return y;
                }
                Debug.WriteLine(")");
            }
            else
            {
                Debug.WriteLine("MIN(");
                if (x < y)
                {
                    Debug.WriteLine(x);
                    return x;
                }
                else
                {
                    Debug.WriteLine(y);
                    return y;
                }
                Debug.WriteLine(")");
            }

        }

        private double WalkLeft(FormExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<FormExcelParser.ExpressionContext>(0));
        }

        private double WalkRight(FormExcelParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<FormExcelParser.ExpressionContext>(1));
        }
    }
}

