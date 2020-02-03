using G1ANT.Language;
using System.Linq;
using System;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.getstate",Tooltip = "This command allows capturing current state displayed in the status bar (left side) at bottom of the form.")]
    public class MsCrmGetStateCommand : Command
    {
        public MsCrmGetStateCommand(AbstractScripter scripter) : base(scripter) { }
        public class Arguments : CommandArguments
        {
            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");       
        }

        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(string.Empty));
            try
            {
                var currentCrm = MsCrmManager.Instance.CurrentCRM;
                string state = currentCrm.GetState();
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(state));
            }
            catch
            {
                throw new ApplicationException("Unable to get state from CRM");
            }
        }
    }
}
