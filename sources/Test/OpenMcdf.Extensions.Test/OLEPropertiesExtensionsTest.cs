﻿using System;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenMcdf.Extensions.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class OlePropertiesExtensionsTest
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
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void Test_SUMMARY_INFO_READ()
        {
            using (var cf = new CompoundFile("_Test.ppt"))
            {
                var co = cf.RootStorage.GetStream("\u0005SummaryInformation").AsOlePropertiesContainer();

                foreach (var p in co.Properties)
                {
                    Debug.Write(p.PropertyName);
                    Debug.Write(" - ");
                    Debug.Write(p.VtType);
                    Debug.Write(" - ");
                    Debug.WriteLine(p.Value);
                }

                Assert.IsNotNull(co.Properties);

                cf.Close();
            }
        }

        [TestMethod]
        public void Test_DOCUMENT_SUMMARY_INFO_READ()
        {
            using (var cf = new CompoundFile("_Test.ppt"))
            {
                var co = cf.RootStorage.GetStream("\u0005DocumentSummaryInformation").AsOlePropertiesContainer();

                foreach (var p in co.Properties)
                {
                    Debug.Write(p.PropertyName);
                    Debug.Write(" - ");
                    Debug.Write(p.VtType);
                    Debug.Write(" - ");
                    Debug.WriteLine(p.Value);
                }

                Assert.IsNotNull(co.Properties);

                cf.Close();
            }

        }

        [TestMethod]
        public void Test_DOCUMENT_SUMMARY_INFO_ROUND_TRIP()
        {
            if (File.Exists("test1.cfs"))
                File.Delete("test1.cfs");

            using (var cf = new CompoundFile("_Test.ppt"))
            {
                var co = cf.RootStorage.GetStream("\u0005DocumentSummaryInformation").AsOlePropertiesContainer();
                using (var cf2 = new CompoundFile())
                {
                    cf2.RootStorage.AddStream("\u0005DocumentSummaryInformation");

                    co.Save(cf2.RootStorage.GetStream("\u0005DocumentSummaryInformation"));

                    cf2.Save("test1.cfs");
                    cf2.Close();
                }

                cf.Close();
            }

        }

        [TestMethod]
        public void Test_SUMMARY_INFO_READ_UTF8_ISSUE_33()
        {
            try
            {
                using (var cf = new CompoundFile("wstr_presets.doc"))
                {
                    var co = cf.RootStorage.GetStream("\u0005SummaryInformation").AsOlePropertiesContainer();

                    foreach (var p in co.Properties)
                    {
                        Debug.Write(p.PropertyName);
                        Debug.Write(" - ");
                        Debug.Write(p.VtType);
                        Debug.Write(" - ");
                        Debug.WriteLine(p.Value);
                    }

                    Assert.IsNotNull(co.Properties);

                    var co2 = cf.RootStorage.GetStream("\u0005DocumentSummaryInformation").AsOlePropertiesContainer();

                    foreach (var p in co2.Properties)
                    {
                        Debug.Write(p.PropertyName);
                        Debug.Write(" - ");
                        Debug.Write(p.VtType);
                        Debug.Write(" - ");
                        Debug.WriteLine(p.Value);
                    }

                    Assert.IsNotNull(co2.Properties);
                }
            }catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Assert.Fail();
            }

        }

        [TestMethod]
        public void Test_SUMMARY_INFO_READ_UTF8_ISSUE_34()
        {
            try
            {
                using (var cf = new CompoundFile("2custom.doc"))
                {
                    var co = cf.RootStorage.GetStream("\u0005SummaryInformation").AsOlePropertiesContainer();

                    foreach (var p in co.Properties)
                    {
                        Debug.Write(p.PropertyName);
                        Debug.Write(" - ");
                        Debug.Write(p.VtType);
                        Debug.Write(" - ");
                        Debug.WriteLine(p.Value);
                    }

                    Assert.IsNotNull(co.Properties);

                    var co2 = cf.RootStorage.GetStream("\u0005DocumentSummaryInformation").AsOlePropertiesContainer();

                    Assert.IsNotNull(co2.Properties);
                    foreach (var p in co2.Properties)
                    {
                        Debug.Write(p.PropertyName);
                        Debug.Write(" - ");
                        Debug.Write(p.VtType);
                        Debug.Write(" - ");
                        Debug.WriteLine(p.Value);
                    }

                    
                    Assert.IsNotNull(co2.UserDefinedProperties.Properties);
                    foreach (var p in co2.UserDefinedProperties.Properties)
                    {
                        Debug.Write(p.PropertyName);
                        Debug.Write(" - ");
                        Debug.Write(p.VtType);
                        Debug.Write(" - ");
                        Debug.WriteLine(p.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Assert.Fail();
            }

        }
        [TestMethod]
        public void Test_SUMMARY_INFO_READ_LPWSTRING()
        {
            using (var cf = new CompoundFile("english.presets.doc"))
            {
                var co = cf.RootStorage.GetStream("\u0005SummaryInformation").AsOlePropertiesContainer();

                foreach (var p in co.Properties)
                {
                    Debug.Write(p.PropertyName);
                    Debug.Write(" - ");
                    Debug.Write(p.VtType);
                    Debug.Write(" - ");
                    Debug.WriteLine(p.Value);
                }

                Assert.IsNotNull(co.Properties);

                cf.Close();
            }

        }

    }
}
