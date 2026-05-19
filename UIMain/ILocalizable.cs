/*
 ILocalizable lets a control re-apply localization that the generic tree-walker
 (Language.LocalizeFrameworkElement) cannot reach: ToolTips, TextBox watermarks,
 DataGrid column headers, Window titles and any code-built labels.

 UIManager.RelocalizeAll() calls LocalizeUI() on every cached control that
 implements this, right after the Tag-based pass, so non-Tag strings refresh
 live when the user switches language (no SpaceClaim restart).

 Implementations must contain ONLY the non-Tag assignments — they must NOT call
 Language.LocalizeFrameworkElement themselves (RelocalizeAll already does that),
 otherwise the logical tree is walked twice.
*/

namespace AESCConstruct2026.Localization
{
    public interface ILocalizable
    {
        void LocalizeUI();
    }
}
