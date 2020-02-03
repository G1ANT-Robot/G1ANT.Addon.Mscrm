using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.previousstage",Tooltip = "This command clicks on 'Previous Stage(Back)' link in CRM. If operation is not possible ♥result  returns false, otherwise returns true.")]
    public class MsCrmPreviousStageCommand : Command
    {
        public MsCrmPreviousStageCommand(AbstractScripter scripter) : base(scripter) { }
        
        public class Arguments : CommandArguments
        {
            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result"); 
        }

        public void Execute(Arguments arguments)
        {
           Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(MsCrmManager.Instance.CurrentCRM.PrevStage()));
        }
    }
}
