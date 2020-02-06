using G1ANT.Language;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace G1ANT.Addon.Mscrm
{
    [Wizard(Name = "Dynamics CRM Recorder", Menu = "Tools\\Wizards")]
    public class MscrmRecorderWizard : Wizard
    {
        public MscrmRecorderWizard() { }

        public override void Execute(AbstractScripter scripter)
        {
            IMainForm mainForm = null;

            foreach (var form in Application.OpenForms)
            {
                if (form is IMainForm)
                {
                    mainForm = form as IMainForm;
                    break;
                }
            }

            using (var recordingWizard = new MscrmRecorderForm(mainForm))
            {
                recordingWizard.ShowDialog();
            }            
        }
    }
}
