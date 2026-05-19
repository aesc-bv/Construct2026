/*
 L is a thin call-site helper over the Language CSV engine.
 - L.T(key)            : translate a key (same as Language.Translate).
 - L.F(key, args)      : translate then string.Format with runtime values ({0}, {1}, ...).
 - L.Status(key, type, args) : translate (+ optional format) and push to the SpaceClaim status bar,
                               removing the StatusMessageType/null boilerplate at every call site.
 Language stays the pure CSV engine; L is only sugar so the ~95 newly-wired sites stay terse.
*/

using SpaceClaim.Api.V242;

namespace AESCConstruct2026.Localization
{
    public static class L
    {
        public static string T(string key) => Language.Translate(key);

        public static string F(string key, params object[] args)
            => (args == null || args.Length == 0)
                ? Language.Translate(key)
                : string.Format(Language.Translate(key), args);

        public static void Status(string key, StatusMessageType type, params object[] args)
            => Application.ReportStatus(F(key, args), type, null);
    }
}
