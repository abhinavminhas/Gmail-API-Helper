using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class TestBase
    {
        private TestContext _testContextInstance;
        protected readonly static string TestEmailId = ConfigurationManager.AppSettings["GmailId"];
        protected readonly string ApplicationName = "GmailAPIHelper";
        protected readonly string EmailDoesNotExistsSearchQuery = $"[from:{TestEmailId}][subject:'Email does not exists.']in:inbox is:unread";

        public TestContext TestContext
        {
            get
            {
                return _testContextInstance;
            }
            set
            {
                _testContextInstance = value;
            }
        }
    }
}