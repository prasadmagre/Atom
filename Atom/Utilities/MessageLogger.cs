using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities
{
    public class MessageLogger
    {
        public static void LogMessage(string loggermessage)
        {
            try
            {
                if (loggermessage != null && loggermessage != "")
                    PageLabelTableManager.GetTestContext().WriteLine(loggermessage);
            }
            catch (Exception e)
            {
                PageLabelTableManager.GetTestContext().WriteLine(loggermessage.Replace('}', ' ').Replace('{', ' ').Trim());
            }
        }
    }
}
