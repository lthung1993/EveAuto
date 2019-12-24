using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Automation.TestFaceBook
{

    public class TestFaceBook : AutomationBase
    {
        [Test]
        public void TestDemoFaceBook()
        {
            driver.Url = "https://www.google.com/";
            _log += "PASS demo Facebook";
            Thread.Sleep(4000);
            Assert.Pass(_log);
        }
    }
}
