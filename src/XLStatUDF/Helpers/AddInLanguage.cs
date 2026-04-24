namespace XLStatUDF.Helpers
{
    public static class AddInLanguage
    {
#if XLSTATUDF_LANG_EN
        public const string Code = "en";
#else
        public const string Code = "cs";
#endif

        public static bool IsCzech => Code == "cs";
    }
}
