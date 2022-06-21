﻿using System.Collections.Generic;
using Expressive.Helpers;

namespace Expressive.Expressions.Binary.Multiplicative
{
    internal class ModulusExpression : BinaryExpressionBase
    {
        #region Constructors

        public ModulusExpression(IExpression lhs, IExpression rhs, Context context) : base(lhs, rhs, context)
        {
        }

        #endregion

        #region BinaryExpressionBase Members

        /// <inheritdoc />
        protected override object EvaluateImpl(object lhsResult, IExpression rightHandSide, IDictionary<string, object> variables) =>
            EvaluateAggregates(lhsResult, rightHandSide, variables, Numbers.Modulus);

        #endregion
    }
}
