using Rubicon.Utilities.Expr;

namespace Bukonf.Gever.Domain
{

    [ExpressionAccess(TypeAccessMode.Full)]
    public static class DeeplExtensions
    {

        public static string Translate(this string text, string quellspache, string zielsprache)
        {
            var result = new DeeplTranslator().Translate(text, quellspache, zielsprache);
            return result;
        }
    }
}
