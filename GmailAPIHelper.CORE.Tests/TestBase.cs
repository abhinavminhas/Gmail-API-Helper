using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class TestBase
    {
        private TestContext _testContextInstance;
        protected readonly string ApplicationName = "GmailAPIHelper";
        protected readonly string EmailDoesNotExistsSearchQuery = "[from:test.auto.helper@gmail.com][subject:'Email does not exists.']in:inbox is:unread";

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