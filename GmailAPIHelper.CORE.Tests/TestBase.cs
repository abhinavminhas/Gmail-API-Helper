using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GmailAPIHelper.CORE.Tests
{
    [TestClass]
    public class TestBase
    {
        private TestContext _testContextInstance;
        protected readonly string ApplicationName = "GmailAPIHelper";

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