using Expressive.Expressions;
using System;

namespace Expressive.Functions.Mathematical
{
    internal class CosFunction : FunctionBase
    {
        #region FunctionBase Members

        public override string Name { get { return "Cos"; } }

        public override object Evaluate(IExpression[] parameters, Context context)
        {
            this.ValidateParameterCount(parameters, 1, 1);
            
            return Math.Cos(Convert.ToDouble(parameters[0].Evaluate(Variables)));
        }

        #endregion
    }
}
