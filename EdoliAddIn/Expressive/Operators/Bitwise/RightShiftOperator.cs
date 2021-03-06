using System.Collections.Generic;
using Expressive.Expressions;
using Expressive.Expressions.Binary.Bitwise;

namespace Expressive.Operators.Bitwise
{
    internal class RightShiftOperator : OperatorBase
    {
        #region OperatorBase Members

        public override IEnumerable<string> Tags => new[] { ">>" };

        public override IExpression BuildExpression(Token previousToken, IExpression[] expressions, Context context)
        {
            return new RightShiftExpression(expressions[0], expressions[1], context);
        }

        public override OperatorPrecedence GetPrecedence(Token previousToken)
        {
            return OperatorPrecedence.RightShift;
        }

        #endregion
    }
}
