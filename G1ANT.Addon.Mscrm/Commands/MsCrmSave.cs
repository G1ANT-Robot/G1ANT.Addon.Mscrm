using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.save",Tooltip = "This command saves changes in current form of CRM.")]
    public class MsCrmSaveCommand : Command
    {
        public MsCrmSaveCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(MsCrmManager.Instance.CurrentCRM.Save(GetTimeoutLeftSeconds())));
        }
    }
}
