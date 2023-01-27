using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenMcdf.Test
{


    /// <summary>
    ///This is a test class for SectorCollectionTest and is intended
    ///to contain all SectorCollectionTest Unit Tests
    ///</summary>
    [TestClass]
    public class SectorCollectionTest
    {


        private TestContext _testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get => _testContextInstance;
            set => _testContextInstance = value;
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for Count
        ///</summary>
        [TestMethod]
        public void CountTest()
        {
            
            var count = 0;

            var target = new SectorCollection();
            int actual;
            actual = target.Count;

            Assert.IsTrue(actual == count);
            var s = new Sector(4096);

            target.Add(s);
            Assert.IsTrue(target.Count == actual + 1);


            for (var i = 0; i < 5000; i++)
                target.Add(s);

            Assert.IsTrue(target.Count == actual + 1 + 5000);
        }

        /// <summary>
        ///A test for Item
        ///</summary>
        [TestMethod]
        public void ItemTest()
        {
            var count = 37;

            var target = new SectorCollection();
            var index = 0;

            var expected = new Sector(4096);
            target.Add(null);

            Sector actual;
            target[index] = expected;
            actual = target[index];

            Assert.AreEqual(expected, actual);
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.Id == expected.Id);

            actual = null;

            try
            {
                actual = target[count + 100];
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfException);
            }

            try
            {
                actual = target[-1];
            }
            catch (Exception ex)
            {
                Assert.IsTrue(ex is CfException);
            }
        }

        /// <summary>
        ///A test for SectorCollection Constructor
        ///</summary>
        [TestMethod]
        public void SectorCollectionConstructorTest()
        {

            var target = new SectorCollection();

            Assert.IsNotNull(target);
            Assert.IsTrue(target.Count == 0);

            var s = new Sector(4096);
            target.Add(s);
            Assert.IsTrue(target.Count == 1);
        }

        /// <summary>
        ///A test for Add
        ///</summary>
        [TestMethod]
        public void AddTest()
        {
            var target = new SectorCollection();
            for (var i = 0; i < 579; i++)
            {
                target.Add(null);
            }


            var item = new Sector(4096);
            target.Add(item);
            Assert.IsTrue(target.Count == 580);
        }

        /// <summary>
        ///A test for GetEnumerator
        ///</summary>
        [TestMethod]
        public void GetEnumeratorTest()
        {
            var target = new SectorCollection();
            for (var i = 0; i < 579; i++)
            {
                target.Add(null);
            }

            
            var item = new Sector(4096);
            target.Add(item);
            Assert.IsTrue(target.Count == 580);

            var cnt = 0;
            foreach (var s in target)
            {
                cnt++;
            }

            Assert.IsTrue(cnt == target.Count);
        }
    }
}
