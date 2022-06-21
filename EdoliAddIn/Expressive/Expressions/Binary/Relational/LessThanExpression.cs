﻿using System.Collections.Generic;
using Expressive.Helpers;

namespace Expressive.Expressions.Binary.Relational
{
    internal class LessThanExpression : BinaryExpressionBase
    {
        #region Constructors

        public LessThanExpression(IExpression lhs, IExpression rhs, Context context) : base(lhs, rhs, context)
        {
        }

        #endregion

        #region BinaryExpressionBase Members

        protected override object EvaluateImpl(object lhsResult, IExpression rightHandSide, IDictionary<string, object> variables) => 
            Comparison.CompareUsingMostPreciseType(lhsResult, rightHandSide.Evaluate(variables), this.Context) < 0;

        #endregion
    }
}
