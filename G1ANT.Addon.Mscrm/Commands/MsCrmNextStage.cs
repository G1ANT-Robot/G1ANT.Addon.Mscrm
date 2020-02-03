using G1ANT.Language;
using System;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.nextstage",Tooltip = "This command clicks on 'Next Stage' link in CRM.")]
    public class MsCrmNextStageCommand : Command
    {
        public MsCrmNextStageCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {  
            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }
        }

        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(MsCrmManager.Instance.CurrentCRM.NextStage()));
        }
    }
}
