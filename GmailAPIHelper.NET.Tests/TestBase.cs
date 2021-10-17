using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GmailAPIHelper.NET.Tests
{
    [TestClass]
    public class TestBase
    {
        private TestContext _testContextInstance;
        protected readonly string ApplicatioName = "GmailAPIHelper";

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