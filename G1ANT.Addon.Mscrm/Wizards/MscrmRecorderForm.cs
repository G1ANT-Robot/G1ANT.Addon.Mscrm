using G1ANT.Language;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace G1ANT.Addon.Mscrm
{
    public partial class MscrmRecorderForm : Form
    {
        private bool stopRecording = false;
        private bool isRecording = false;
        private Thread recordingThread = null;
        private int xPosition;
        private int yPosition;
        IMainForm mainForm;

        public MscrmRecorderForm(IMainForm mainForm)
        {
            this.mainForm = mainForm;
            InitializeComponent();
            SetUrlFromRegistry();
        }

        private void SetUrlInRegistry(string value)
        {
            string registryLocation = @"HKEY_CURRENT_USER\Software\G1ant\Platform";
            string registryKey = "MsCrmWizardUrl";
            Registry.SetValue(registryLocation, registryKey, value);
        }

        private void SetUrlFromRegistry()
        {
            string registryLocation = @"HKEY_CURRENT_USER\Software\G1ant\Platform";
            string registryKey = "MsCrmWizardUrl";
            string value = Registry.GetValue(registryLocation, registryKey, null) as string;
            if (string.IsNullOrEmpty(value) == false)
            {
                UrlPhraseTextBox.Text = value;
            }
        }

        private void StartRecordingButton_Click(object sender, EventArgs e)
        {
            if (!isRecording)
            {
                recordingThread = new Thread(StartRecording)
                {
                    IsBackground = true
                };
                stopRecording = false;
                recordingThread.SetApartmentState(ApartmentState.STA);
                recordingThread.Start();
                isRecording = true;
                ((Button)sender).Text = "Stop recording";
            }
            else
            {
                if (recordingThread != null && recordingThread.IsAlive)
                {
                    stopRecording = true;
                }
                while (recordingThread != null && recordingThread.IsAlive)
                {
                    // because recordingThread is STA thread, recordingThread.Join() prevented recordingThread from exiting
                    Thread.Sleep(10);
                    Application.DoEvents();
                }
                stopRecording = false;
                isRecording = false;
                ((Button)sender).Text = "Start recording";
            }
        }

        private void StartRecording()
        {
            bool justStarted = true;
            while (!stopRecording)
            {
                try
                {
                    var wrapper = MsCrmManager.Instance.AttachToExistingCRM(MsCrmUrlPhrase, "url", true);
                    if (wrapper.Ie == null && justStarted == true)
                    {
                        MessageBox.Show($"Dynamics CRM Instance with url containing phrase '{MsCrmUrlPhrase}' could not be found.",
                                        "Dynamics CRM Instance not found",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                        isRecording = false;
                        StartRecordingButton.Invoke((MethodInvoker)(() =>
                        {
                            StartRecordingButton.Text = "Start recording";
                        }));
                        return;
                    }
                    else
                        justStarted = false;
                    wrapper.Ie.ShowWindow(WatiN.Core.Native.Windows.NativeMethods.WindowShowStyle.Minimize);
                    wrapper.Ie.ShowWindow(WatiN.Core.Native.Windows.NativeMethods.WindowShowStyle.Restore);
                    wrapper.Ie.ShowWindow(WatiN.Core.Native.Windows.NativeMethods.WindowShowStyle.Restore);

                    wrapper.Ie.RunScript("$('.G1antMsCrmWizardClass').remove()", "javascript");
                    while (!stopRecording)
                    {
                        var div = MsCrmManager.Instance.CurrentCRM.Ie.Div("MsCrmWizard");
                        while (!stopRecording && !div.Exists)
                        {
                            try
                            {
                                div.WaitUntilExists(1);
                            }
                            catch { }
                            Thread.Sleep(100);
                        }
                        if (!stopRecording)
                        {
                            var action = div.GetAttributeValue("data-element-type");
                            string searchBy = div.GetAttributeValue("target_by");
                            string searchPhrase = div.GetAttributeValue("target_search");
                            string command = string.Empty;
                            bool addCommand = true;
                            switch (action)
                            {
                                case "click":
                                    command = searchBy == "id" ? $"mscrm.click ‴{searchPhrase}‴" :
                                                                 $"mscrm.click by ‴{searchBy}‴ search ‴{searchPhrase}‴";
                                    break;
                                case "setvalue":                                    
                                    Command.ExecutionTime = Environment.TickCount;
                                    Command.Timeout = 10000;
                                    var element = wrapper.FindElement(searchPhrase, searchBy);
                                    List<string> availableOptions = null;
                                    var fieldType = wrapper.RecognizeFieldType(element, ref availableOptions);
                                    MsCrmSetValueForm setValueForm = new MsCrmSetValueForm(searchPhrase, availableOptions);
                                    wrapper.TypeText("⋘esc⋙");
                                    setValueForm.StartPosition = FormStartPosition.Manual;
                                    setValueForm.Locate(xPosition, yPosition);
                                    setValueForm.ShowDialog();
                                    if (setValueForm.DialogResult == DialogResult.OK)
                                    {
                                        switch (setValueForm.ResultType)
                                        {
                                            case ResultType.SetValue:
                                                command = searchBy == "id" ? $"mscrm.setvalue ‴{searchPhrase}‴ value ‴{setValueForm.ResultValue}‴" :
                                                $"mscrm.setvalue by ‴{searchBy}‴ search ‴{searchPhrase}‴ value ‴{setValueForm.ResultValue}‴";

                                                if (setValueForm.ValidationErrorMessage != null)
                                                {
                                                    command += $@"{System.Environment.NewLine}dialog ‴{setValueForm.ValidationErrorMessage}‴ if {SpecialChars.Variable}validationresult==false{System.Environment.NewLine}stop if {SpecialChars.Variable}validationresult==false";
                                                }
                                                Command.ExecutionTime = Environment.TickCount;
                                                Command.Timeout = 10000;
                                                try
                                                {
                                                    wrapper.SetVal(searchPhrase, setValueForm.ResultValue, searchBy);
                                                }
                                                catch { }
                                                break;
                                            case ResultType.Click:
                                                command = searchBy == "id" ? $"mscrm.click ‴{searchPhrase}‴" :
                                                                 $"mscrm.click by ‴{searchBy}‴ search ‴{searchPhrase}‴";
                                                break;
                                        }
                                    }
                                    else
                                        addCommand = false;
                                    break;
                            }
                            if (addCommand)
                                AppendRecordingRichTextBox(command);
                            wrapper.Ie.RunScript("$('.G1antMsCrmWizardClass').remove()", "javascript");
                        }
                    }
                    try
                    {
                        wrapper.MsCrmRecordingSwitch = false;
                        wrapper.Detach(true);
                    }
                    catch { }
                }
                catch { }
            }
            stopRecording = false;
            isRecording = false;
        }

        private void AppendRecordingRichTextBox(string content)
        {
            RecordedScriptRichTextBox.Invoke((MethodInvoker)(() =>
            {
                if (RecordedScriptRichTextBox.Text.EndsWith($"{content}\n") == false)
                {
                    RecordedScriptRichTextBox.AppendText($"{content}\n");
                    RecordedScriptRichTextBox.SelectionStart = RecordedScriptRichTextBox.Text.Length;
                    RecordedScriptRichTextBox.ScrollToCaret();
                }
            }));
        }

        private string MsCrmUrlPhrase
        {
            get
            {
                var returnValue = string.Empty;
                UrlPhraseTextBox.Invoke((MethodInvoker)(() => { returnValue = UrlPhraseTextBox.Text; }));
                return returnValue;
            }
        }

        private void MscrmRecorderForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            stopRecording = true;
            StartRecordingButton.Text = "Start recording";
            //e.Cancel = true;
            //this.Hide();
        }

        private void InsertScriptToMainWindow_Click(object sender, EventArgs e)
        {
            string textToInsert = AddMscrmAttachIntroCheckBox.Checked ?
                $"mscrm.attach ‴{MsCrmUrlPhrase}‴{System.Environment.NewLine}{RecordedScriptRichTextBox.Text}" :
                RecordedScriptRichTextBox.Text;

            mainForm?.InsertTextIntoCurrentEditor(textToInsert);
        }

        private void UrlPhraseTextBox_TextChanged(object sender, EventArgs e)
        {
            SetUrlInRegistry(((TextBox)sender).Text);
        }

        public void LocateInMiddleOfForm(Form form)
        {
            var x = form.Location.X + 50;
            var y = form.Location.Y + 50;
            Location = new Point(x, y);
        }

        private void MscrmRecorderForm_LocationChanged(object sender, EventArgs e)
        {
            if (Location.X > 0)
                xPosition = Location.X;
            if (Location.Y > 0)
                yPosition = Location.Y;
        }
    }
}
