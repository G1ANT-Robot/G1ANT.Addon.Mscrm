using System;
using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.detach",Tooltip = "This command disconnect from instance of CRM attached by 'mscrm.attach'.")]
    class MsCrmDetachCommand : Command
    {
        public MsCrmDetachCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result"); 
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                if (MsCrmManager.Instance.CurrentCRM != null)
                {
                    MsCrmManager.Instance.Detach(MsCrmManager.Instance.CurrentCRM);
                }
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(true));
            }
            catch
            {
                throw new ApplicationException("Unable to Detach from CRM");
            }
        }
    }
}
