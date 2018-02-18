using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities.TestDataExcel
{
    public class VariableManager
    {
        /// <summary>
        /// Method to get value from variable
        /// </summary>
        /// <param name="variableName"></param>
        /// <returns></returns>
        public static string GetVariableValue(string variableName)
        {
            variableName = variableName.Replace('{', ' ').Replace('}', ' ').Trim();
            string variableValue = "";
            if (variableName.ToUpper().StartsWith("G_"))
            {
                //variableValue = TestDataProvider.allVariable["G"][variableName.ToUpper()];
            }
            else if (variableName.ToUpper().StartsWith("M_"))
            {
                string moduleIdentifier = variableName.Split('_')[1];
               // variableValue = TestDataProvider.allVariable[moduleIdentifier][variableName.ToUpper()];
            }
            else
            {
                variableValue = variableName;


            }

            return variableValue;
        }

        /// <summary>
        /// To check whether we have varible in dictionary
        /// </summary>
        /// <param name="variableName"></param>
        /// <returns></returns>
        public static bool ContainsVariableValue(string variableName)
        {
            bool isContains = false;
            variableName = variableName.Replace('{', ' ').Replace('}', ' ').Trim();

            if (variableName.ToUpper().StartsWith("G_"))
            {
                //isContains = TestDataProvider.allVariable["G"].ContainsKey(variableName);
            }
            else if (variableName.ToUpper().StartsWith("M_"))
            {
                string moduleIdentifier = variableName.Split('_')[1];
                //isContains = TestDataProvider.allVariable[moduleIdentifier].ContainsKey(variableName);
            }


            return isContains;

        }
    }
}
