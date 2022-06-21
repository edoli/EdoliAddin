﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Expressive.Helpers;

namespace Expressive.Expressions.Unary.Additive
{
    internal class MinusExpression : UnaryExpressionBase
    {
        #region Constructors

        public MinusExpression(IExpression expression) : base(expression)
        {
        }

        #endregion

        #region BinaryExpressionBase Members

        public override object Evaluate(IDictionary<string, object> variables) =>
            Numbers.Subtract(0, this.expression.Evaluate(variables));

        #endregion
    }
}
