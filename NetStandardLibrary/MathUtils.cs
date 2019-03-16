using System;
using System.Runtime.InteropServices;

namespace NetStandardLibrary
{
    [Guid("45f838c9-d8cd-4198-b2b3-f4e8c2a5b588")]
    [ComVisible(true)]
    public interface IMathUtils
    {
        double CalculateComplexMethod();
        void CommandWithParameters(int number1, int number2);
    }

    [Guid("293bcd3a-e771-45d5-8a53-5413997c2de8")]
    [ComVisible(true)]
    [ProgId("NetStandardLibrary.MathUtils")]
    public class MathUtils : IMathUtils
    {
        // For COM interop we need parameterless constructor
        public MathUtils() { }

        public double CalculateComplexMethod() => 2 * 2;

        public void CommandWithParameters(int number1, int number2)
        {
            // Do something in here like saving to a JSON file, POST to some REST API or anything
            // without needing to code in prehistoric VB6
        }
    }
}