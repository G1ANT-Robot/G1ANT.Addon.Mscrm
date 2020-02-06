using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.finish",Tooltip = "This command clicks on 'Finish' button in CRM.")]
    public class MsCrmFinishCommand : Command
    {
        public MsCrmFinishCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }
        }

        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(MsCrmManager.Instance.CurrentCRM.Finish()));
        }
    }
}
