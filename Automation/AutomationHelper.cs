using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Automation
{
    public class AutomationHelper
    {
        public static int DefaultTimeOutSeconds
        {
            get
            {
               return int.Parse(System.Configuration.ConfigurationSettings.AppSettings["DefaultTimeOutSeconds"]); 
            }
        }
        public static IWebElement WaitElement(IWebDriver webDriver, By by, int timeoutSeconds = 0, IWebElement container = null)
        {
            if (timeoutSeconds == 0) timeoutSeconds = DefaultTimeOutSeconds;
            try
            {
                WebDriverWait waiter = new WebDriverWait(webDriver, TimeSpan.FromSeconds(timeoutSeconds));

                return waiter.Until((d) =>
                {
                    try
                    {
                        if (container != null)
                        {
                            return container.FindElement(by);
                        }
                        else
                        {
                            return d.FindElement(by);
                        }
                    }
                    catch (Exception)
                    {
                        return null;
                    }
                });
            }
            catch
            {
                return null;
            }
        }
        public static ReadOnlyCollection<IWebElement> WaitElements(IWebDriver webDriver, By by, bool allowEmptyList = false, int timeoutSeconds = 0, IWebElement container = null)
        {
            if (timeoutSeconds == 0) timeoutSeconds = DefaultTimeOutSeconds;
            try
            {
                WebDriverWait waiter = new WebDriverWait(webDriver, TimeSpan.FromSeconds(timeoutSeconds));
                return waiter.Until((d) =>
                {
                    try
                    {
                        if (container != null)
                        {
                            var tmp = container.FindElements(by);
                            if (!tmp.Any() && !allowEmptyList) return null;
                            return tmp;
                        }
                        else
                        {
                            var tmp = d.FindElements(by);
                            if (!tmp.Any() && !allowEmptyList) return null;
                            return tmp;
                        }
                    }
                    catch (Exception)
                    {

                        return null;
                    }
                });
            }
            catch
            {
                return null;
            }
        }
    }
}
