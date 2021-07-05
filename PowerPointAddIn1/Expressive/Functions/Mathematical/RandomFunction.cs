﻿using Expressive.Expressions;
using System;

namespace Expressive.Functions.Mathematical
{
    internal class RandomFunction : FunctionBase
    {
        #region FunctionBase Members

        public override string Name { get { return "Random"; } }

        public override object Evaluate(IExpression[] parameters, Context context)
        {
            this.ValidateParameterCount(parameters, 2, 2);

            object min = parameters[0].Evaluate(Variables);
            object max = parameters[1].Evaluate(Variables);
            
            var random = new Random(DateTime.UtcNow.Millisecond);

            if (min is int && max is int)
            {
                return random.Next((int)min, (int)max);
            }
            else if (min is double || max is double)
            {
                var value = random.NextDouble();
                var typedMin = Convert.ToDouble(min);
                var scale = Convert.ToDouble(max) - typedMin;

                return typedMin + (scale * value);
            }

            return null;
        }

        #endregion
    }
}
