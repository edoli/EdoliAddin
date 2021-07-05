﻿using Expressive.Expressions;
using System;

namespace Expressive.Functions.Conversion
{
    internal sealed class IntegerFunction : FunctionBase
    {
        #region FunctionBase Members

        public override string Name => "Integer";

        public override object Evaluate(IExpression[] parameters, Context context)
        {
            this.ValidateParameterCount(parameters, 1, 1);

            var objectToConvert = parameters[0].Evaluate(this.Variables);

            // No point converting if there is nothing to convert.
            if (objectToConvert is null) { return null; }

            return Convert.ToInt32(objectToConvert, context.CurrentCulture);
        }

        #endregion
    }
}
