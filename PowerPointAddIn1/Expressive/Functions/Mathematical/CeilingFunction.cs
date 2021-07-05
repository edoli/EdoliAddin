﻿using Expressive.Expressions;
using System;

namespace Expressive.Functions.Mathematical
{
    internal class CeilingFunction : FunctionBase
    {
        #region FunctionBase Members

        public override string Name { get { return "Ceiling"; } }

        public override object Evaluate(IExpression[] parameters, Context context)
        {
            this.ValidateParameterCount(parameters, 1, 1);

            var value = parameters[0].Evaluate(Variables);

            if (value is double)
            {
                return Math.Ceiling((double)value);
            }
            else if (value is decimal)
            {
                return Math.Ceiling((decimal)value);
            }
            return Math.Ceiling(Convert.ToDouble(value));
        }

        #endregion
    }
}
