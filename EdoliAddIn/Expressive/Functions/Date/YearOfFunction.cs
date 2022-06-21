﻿using Expressive.Expressions;
using System;

namespace Expressive.Functions.Date
{
    internal sealed class YearOfFunction : FunctionBase
    {
        #region FunctionBase Members

        public override string Name => "YearOf";

        public override object Evaluate(IExpression[] parameters, Context context)
        {
            this.ValidateParameterCount(parameters, 1, 1);

            var dateObject = parameters[0].Evaluate(this.Variables);

            if (dateObject is null) { return null; }

            var date = Convert.ToDateTime(dateObject, context.CurrentCulture);

            return date.Year;
        }

        #endregion
    }
}
