using G1ANT.Addon.MsCrm.Properties;
using mshtml;
using System;
using System.Reflection;
using System.Runtime.InteropServices.Expando;
using WatiN.Core;
using WatiN.Core.Native;
using WatiN.Core.Native.InternetExplorer;

namespace G1ANT.Addon.Mscrm
{
    public static class WatiNHelper
    {
        public static string MsCrmFindViaJquerySelector
        {
            get
            {
                return Resources.FindViaJquerySelector;
            }
        }

        public static string AreCommandBarsInitialized
        {
            get
            {
                return Resources.AreCommandBarsInitialized.Replace("\r", "").Replace("\n", "");
            }
        }

        public static T Eval<T>(this INativeDocument document, string code, string type)
        {
            IExpando window = GetWindow(document);
            PropertyInfo property = CreateOrGetProperty(window, "__lastEvalResult");
            document.RunScript("window.__lastEvalResult = " + code + ";", type);
            var result = (T)property.GetValue(window, null);
            return result;
        }

        private static PropertyInfo CreateOrGetProperty(IExpando expando, string name)
        {
            var currentProperty = expando.GetProperty(name, BindingFlags.GetProperty);
            return currentProperty ?? expando.AddProperty(name);
        }

        private static void RemoveProperty(IExpando expando, PropertyInfo property)
        {
            expando.RemoveMember(property);
        }

        private static IExpando GetWindow(INativeDocument document)
        {
            IEDocument ieDoc = document as IEDocument;
            IHTMLDocument2 htmlDoc = ieDoc.HtmlDocument;
            return htmlDoc.parentWindow as IExpando;
        }

        public static ElementCollection Elements(this Element self)
        {
            return (self as List)?.Elements ??
                   (self as Div)?.Elements ??
                   (self as ListItem)?.Elements ??
                   (self as ElementContainer<ListItem>)?.Elements;
        }

        public static bool IsJavascriptEvaluateTrue(this Browser browser, string javascript)
        {
            var evalResult = browser.Eval("if (" + javascript + ") {true;} else {false;}");
            return evalResult == "true";
        }

        public static string FindIdViaJQuery(this Browser browser, string cssSelector, string iFrame, int timeoutMs)
        {
            int start = Environment.TickCount;
            while (Environment.TickCount - start <= timeoutMs)
            {
                if (!browser.IsJavascriptEvaluateTrue("typeof('WatinSearchHelper') == 'function'"))
                {
                    var js = MsCrmFindViaJquerySelector.Replace("\n", "").Replace("\r", "");
                    browser.Eval(js);
                }
                var elementId = browser.Eval(string.Format("WatinSearchHelper.getElementId(\"{0}\", \"{1}\")", cssSelector, "none"));

                if (elementId != "_no_element_" && !string.IsNullOrEmpty(elementId))
                    return elementId;
            }
            throw new ArgumentException("Unable to find element on page with selector '" + cssSelector + "'", "cssSelector");
        }
    }
}
