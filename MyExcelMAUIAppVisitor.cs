using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Linq.Expressions;


namespace MyExcelMAUIApp
{
    class MyExcelMAUIAppVisitor : MyExcelMAUIAppBaseVisitor<double>
    {
        public static bool IsNumberExpression(string expression)
        {
            string numberPattern = @"^[+-]?\d+(\.\d+)?$";
            return Regex.IsMatch(expression, numberPattern);
        }
        private static bool IsCellReference(string expression)
        {
            string cellReferencePattern = @"^[A-Za-z]+\d+$";
            return Regex.IsMatch(expression, cellReferencePattern);
        }

        static private Dictionary<string, double> variableTable = new Dictionary<string, double>();
        private bool containsComparisonOperation = false;

        public override double VisitCompileUnit(MyExcelMAUIAppParser.CompileUnitContext context)
        {
            var result = Visit(context.expression());
            if (!containsComparisonOperation)
            {
                throw new ArgumentException("Invalid expression!");
                
                
            }
            return result;
        }

        public override double VisitNumberExpr(MyExcelMAUIAppParser.NumberExprContext context)
        {
            return double.Parse(context.GetText());
        }

        public override double VisitIdentifierExpr(MyExcelMAUIAppParser.IdentifierExprContext context)
        {
            var variableName = context.GetText();
            if (variableTable.TryGetValue(variableName, out var value))
            {
                return value;
            }
            else
            {
                return double.NaN;
            }
        }

        public override double VisitParenthesizedExpr(MyExcelMAUIAppParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitAdditiveExpr(MyExcelMAUIAppParser.AdditiveExprContext context)
        {
            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));

            if (context.operatorToken.Type == MyExcelMAUIAppLexer.ADD)
            {
                return left + right;
            }
            else // MyExcelMAUIAppLexer.SUBTRACT
            {
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(MyExcelMAUIAppParser.MultiplicativeExprContext context)
        {
            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));

            if (context.operatorToken.Type == MyExcelMAUIAppLexer.MULTIPLY)
            {
                return left * right;
            }
            else // MyExcelMAUIAppLexer.DIVIDE
            {
                if (right == 0)
                {
                    throw new DivideByZeroException("Division by zero is not allowed.");
                }
                return left / right;
            }
        }

        public override double VisitEqualityExpr(MyExcelMAUIAppParser.EqualityExprContext context)
        {
            containsComparisonOperation = true;

            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));
            return left == right ? 1.0 : 0.0;
        }

        public override double VisitLessThanExpr(MyExcelMAUIAppParser.LessThanExprContext context)
        {
            containsComparisonOperation = true;

            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));
            return left < right ? 1.0 : 0.0;
        }

        public override double VisitGreaterThanExpr(MyExcelMAUIAppParser.GreaterThanExprContext context)
        {
            containsComparisonOperation = true;

            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));
            return left > right ? 1.0 : 0.0;
        }

        public override double VisitNotExpr(MyExcelMAUIAppParser.NotExprContext context)
        {
            var operand = Visit(context.expression());
            return operand == 0.0 ? 1.0 : 0.0;
        }

        public override double VisitAndExpr(MyExcelMAUIAppParser.AndExprContext context)
        {
            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));
            return (left != 0.0 && right != 0.0) ? 1.0 : 0.0;
        }

        public override double VisitOrExpr(MyExcelMAUIAppParser.OrExprContext context)
        {
            var left = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(0));
            var right = Visit(context.GetRuleContext<MyExcelMAUIAppParser.ExpressionContext>(1));
            return (left != 0.0 || right != 0.0) ? 1.0 : 0.0;
        }

        public override double VisitIncrementExpr(MyExcelMAUIAppParser.IncrementExprContext context)
        {
            var val = Visit(context.expression());
            return val + 1;
        }

        public override double VisitDecrementExpr(MyExcelMAUIAppParser.DecrementExprContext context)
        {
            var val = Visit(context.expression());
            return val - 1;
        }

        public static void SetVariable(string variableName, double value)
        {
            variableTable[variableName] = value;
        }
    }
}
