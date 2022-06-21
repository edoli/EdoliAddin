﻿using System;
using System.Collections.Generic;

namespace Expressive.Expressions.Binary.Logical
{
    internal class OrExpression : BinaryExpressionBase
    {
        #region Constructors

        public OrExpression(IExpression lhs, IExpression rhs, Context context) : base(lhs, rhs, context)
        {
        }

        #endregion

        #region BinaryExpressionBase Members

        protected override object EvaluateImpl(object lhsResult, IExpression rightHandSide, IDictionary<string, object> variables) =>
            Convert.ToBoolean(lhsResult) ||
            Convert.ToBoolean(rightHandSide.Evaluate(variables));

        #endregion
    }
}
