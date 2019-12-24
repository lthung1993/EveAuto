using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Automation.TestGoogle
{
    public class TestGoogle : AutomationBase
    {

        [Test]
        public void TestDemo()
        {
            driver.Url = "https://www.google.com/";
            _log += "PASS demo Google";
            Thread.Sleep(4000);
            Assert.Pass(_log);
        }
    }
}
