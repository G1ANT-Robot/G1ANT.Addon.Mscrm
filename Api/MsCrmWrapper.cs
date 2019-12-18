using G1ANT.Addon.MsCrm.Properties;
using G1ANT.Language;
using System;
using System.Collections.Generic;
using System.Threading;
using WatiN.Core;
using G1ANT.Addon.Mscrm.Api;

namespace G1ANT.Addon.Mscrm
{
    public class MsCrmWrapper
    {
        public int Id { get; set; }

        public IE Ie { get; private set; }

        public string Title
        {
            get { return Ie == null ? "" : Ie.Title; }
        }

        public bool MsCrmRecordingSwitch { get; set; } = false;

        public string Address { get; set; }
        public MsCrmWrapper(int id)
        {
            this.Id = id;
        }

        private string MsCrmWizardInjectionScript
        {
            get
            {
                return Resources.MsCrmWizardInjection;
            }
        }

        public MsCrmWrapper(int id, string name, string by = "title", bool mscrmRecorder = false)
        {
            Ie = null;
            try
            {
                this.Id = id;
                string search = name.ToLower();
                IECollection list = IE.InternetExplorers();
                foreach (var item in list)
                {
                    if (by.ToLower() == "title")
                    {
                        if (item.Title.ToLower().Contains(search))
                        {
                            Ie = item;
                            break;
                        }
                    }
                    else
                    {
                        if (item.Url.ToLower().Contains(search))
                        {
                            Ie = item;
                            break;
                        }
                    }
                }

                if (Ie != null)
                {
                    Address = Ie.Url;
                }
            }
            catch
            {
                throw new TimeoutException("Timeout occured while attaching to Internet Explorer / CRM");
            }
            if (mscrmRecorder && Ie?.InternetExplorer != null)
            {
                MsCrmRecordingSwitch = true;
                SHDocVw.InternetExplorer webBrowser = (SHDocVw.InternetExplorer)Ie.InternetExplorer;
                webBrowser.DocumentComplete += ReinjectJavascriptListeners;
                object unimportant = null;
                ReinjectJavascriptListeners(null, ref unimportant);
            }
        }

        private void ReinjectJavascriptListeners(object pDisp, ref object URL)
        {
            {
                var injectionThread = new Thread(() =>
                {
                    string javaScriptCode = MsCrmWizardInjectionScript;
                    try
                    {
                        Ie.NativeDocument.RunScript(javaScriptCode, "javascript");
                    }
                    catch { }
                })
                { IsBackground = true };
                injectionThread.SetApartmentState(ApartmentState.STA);
                injectionThread.Start();
            }
        }

        public void TypeText(string text)
        {
            KeyboardTyper.TypeWithSendInput(text, string.Empty, Ie.hWnd, IntPtr.Zero, 40000, true, 10);
        }

        public void ActivateTab(string by, string phrase)
        {
            if (Ie != null)
            {
                new TabActivator(Ie.hWnd).ActivateTab(by, phrase);
            }
        }

        public void ClickByElement(string search, string by = "id", bool trigger = false, bool noWait = false)
        {
            Element continueLink = FindElement(search, by);
            if (continueLink == null)
            {
                throw new ApplicationException("Element to click not found");
            }
            
            if (trigger)
            {
                if (noWait)
                {
                    continueLink?.FireEventNoWait("onclick");
                }
                else
                {
                    continueLink?.FireEvent("onclick");
                }
            }

            if (!noWait)
            {
                continueLink?.WaitUntilExists();
                continueLink?.WaitForComplete();
                continueLink?.Click();
            }
            else if (noWait)
            {
                continueLink?.ClickNoWait();
            }
        }

        public Element FindElement(string search, string by)
        {
            Element element = null;
            switch(by)
            {
                case "name":
                    element = Ie.Element(Find.ByName(search));
                    break;
                case "title":
                    element = Ie.Element(Find.ByTitle(search));
                    break;
                case "role":
                    element = Ie.Element(Find.By("role", search));
                    break;
                case "selector":
                    element = Ie.Element(Find.ById(WatiNHelper.FindIdViaJQuery(Browser.AttachTo<IE>(Find.ByTitle(Ie.Url)), search, "", Command.Timeout)));
                    break;
                default:
                    element = Ie.Element(Find.ById(search));
                    break;
            }
            return element;
        }

        public FieldType RecognizeFieldType(Element element, ref List<string> availableOptions)
        {
            FieldType result = 0;
            if(element.ClassName.Contains("header_process_parentcontactid.fieldControl-LookupResultsDropdown_parentcontactid_textInputBox_with_filter_new"))
            {
                result = FieldType.Lookup;
            }
            else if (element.TagName.ToLower().Contains("select"))
            {
                result = FieldType.Select;
            }
            else if (element.GetAttributeValue("type").Contains("text") || element.TagName == "textarea")
            {
                result = FieldType.TextField;
            }
            return result;
        }

        public void SetVal(string search, string value, string by = "id")
        {
            Element element = FindElement(search, by);
            element.WaitForComplete();
            List<string> availableOptions = null;
            var fieldType = RecognizeFieldType(element, ref availableOptions);
            switch (fieldType)
            {
                case FieldType.Lookup:
                    SetValInLookupField(element, value);
                    break;

                case FieldType.Select:
                    SetValInSelectItem(element, value);
                    break;

                case FieldType.TextField:
                    SetValInTextField(element, value);
                    break;

                case FieldType.DateTimeTextField:
                    SetValInDateTimeTextField(element, value);
                    break;
            }
        }

        private void SetValInDateTimeTextField(Element element, string value)
        {
            element.Focus();
            element.Click();
            TypeText(value);
            TypeText("⋘ENTER⋙");
        }

        private void SetValInTextField(Element element, string value)
        {
            element.Focus();
            element.Click();
            TypeText(value);
        }

        private void SetValInLookupField(Element element, string value)
        {
            element.Focus();
            element.Click();
            TypeText(value);
            TypeText("⋘DOWN⋙");
            TypeText("⋘ENTER⋙");
        }

        private void SetValInSelectItem(Element element, string value)
        {
            element.Focus();
            TypeText(value);
        }

        public void WaitForLoad(string search, string by = "id", int timeout = 10000)
        {
            Ie.WaitForComplete();
            Element element = FindElement(search, by);
            if (element != null)
            {
                element.WaitUntilExists(timeout / 1000 / 2);
                WatiN.Core.Settings.WaitForCompleteTimeOut = timeout / 2;
                element.WaitForComplete();
            }
            else
            {
                throw new ApplicationException("Element not found");
            }
        }

        public void Detach(bool stopMsCrmRecorder = false)
        {
            if (Ie != null)
            {
                Ie.AutoClose = false;
                Ie.Dispose();
            }
        }


        public bool IsElementVisible(string id, string by = "id")
        {
            Element element = FindElement(id, by);
            bool visible = element.Style.GetAttributeValue("display") != "none" &&
                           element.Style.GetAttributeValue("visibility") != "hidden";
            while (element != null && visible == true)
            {
                element = element.Parent;
                if (element != null)
                {
                    visible = element.Style.GetAttributeValue("display") != "none" &&
                              element.Style.GetAttributeValue("visibility") != "hidden";
                }
            }
            return visible;
        }

        public string GetState()
        {
            var status = FindElement("status", "role");
            return status.InnerHtml ?? throw new Exception("Failed to get status!"); 
        }

        public bool Save(int timeout)
        {
            bool result;
            try
            {
                Ie.Element(Find.ByTitle("Save")).Click();
                Thread.Sleep(300);
                Ie.Element(Find.ByTitle("saving")).WaitUntilRemoved(timeout);
                Ie.WaitForComplete();
                result = true;
            }
            catch
            {
                throw new ApplicationException("Unable to save");
            }
            return result;
        }

        public bool NextStage()
        {
            try
            {
                var result = false;
                var nextStageButton = Ie.Element(Find.ById("MscrmControls.Containers.ProcessStageControl-nextButtonContainer"));
                if (nextStageButton.Exists)
                {
                    nextStageButton.Click();
                    Ie.WaitForComplete();
                    result = true;
                }
                else
                {
                    result = false;
                }
                return result;
            }
            catch 
            {
                throw new Exception("Unable to proceed to the next stage!");
            }
        }

        public bool PrevStage()
        {
            try
            {
                var result = false;
                var prevStageButton = Ie.Element(Find.ById("MscrmControls.Containers.ProcessStageControl-previousButtonContainer"));
                if (prevStageButton.Exists)
                {
                    prevStageButton.Click();
                    Ie.WaitForComplete();
                    result = true;
                }
                else
                {
                    result = false;
                }
                return result;
            }
            catch
            {
                throw new Exception("Unable to proceed to the next stage!");
            }
        }

        public bool Finish()
        {
            try
            {
                var result = false;
                var finishButton = Ie.Element(Find.ById("MscrmControls.Containers.ProcessStageControl-finishButtonContainerbuttonInnerContainer"));
                if(finishButton.Exists)
                {
                    finishButton.Click();
                    Ie.WaitForComplete();
                    result = true;
                }
                else
                {
                    result = false;
                }
                return result;
            }
            catch
            {
                throw new Exception("Unable to Finish the form! Check whether previous stages are completed");
            }
        }

        public string GetValue(string search, string by)
        {
            try
            {
                string value = null;
                var element = FindElement(search, by);
                if(element.Exists)
                {
                    value = element.InnerHtml;
                }
                return value;
            }
            catch
            {
                throw new Exception("Unable to retrieve value for field");
            }
        }
    }
}
