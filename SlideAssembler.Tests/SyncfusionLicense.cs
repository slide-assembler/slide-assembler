namespace SlideAssembler.Tests
{
    [TestClass]
    static class SyncfusionLicense
    {
        [AssemblyInitialize]
        public static void Register(TestContext _)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(Environment.GetEnvironmentVariable("SyncfusionLicenseKey");
        }
    }
}
