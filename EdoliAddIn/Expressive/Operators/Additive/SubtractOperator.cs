﻿using Expressive.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using Expressive.Expressions.Binary.Additive;
using Expressive.Expressions.Unary.Additive;

namespace Expressive.Operators.Additive
{
    internal class SubtractOperator : OperatorBase
    {
        #region IOperator Members

        public override IEnumerable<string> Tags => new[] { "-", "\u2212" };

        public override IExpression BuildExpression(Token previousToken, IExpression[] expressions, Context context)
        {
            if (IsUnary(previousToken))
            {
                return new MinusExpression(expressions[0] ?? expressions[1]);
            }

            return new SubtractExpression(expressions[0], expressions[1], context);
        }

        public override bool CanGetCaptiveTokens(Token previousToken, Token token, Queue<Token> remainingTokens)
        {
            var remainingTokensCopy = new Queue<Token>(remainingTokens.ToArray());

            return this.GetCaptiveTokens(previousToken, token, remainingTokensCopy).Any();
        }

        public override Token[] GetInnerCaptiveTokens(Token[] allCaptiveTokens)
        {
            return allCaptiveTokens.Skip(1).ToArray();
        }

        public override OperatorPrecedence GetPrecedence(Token previousToken)
        {
            return IsUnary(previousToken) ? OperatorPrecedence.UnaryMinus : OperatorPrecedence.Subtract;
        }

        #endregion

        private static bool IsUnary(Token previousToken)
        {
            return string.IsNullOrEmpty(previousToken?.CurrentToken) ||
                string.Equals(previousToken.CurrentToken, "(", StringComparison.Ordinal) ||
                previousToken.CurrentToken.IsArithmeticOperator();
        }
    }
}
