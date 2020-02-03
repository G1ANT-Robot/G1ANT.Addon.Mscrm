using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.getvalue", Tooltip = "This command allows getting the value of the specified field. Make sure the 'Search' Value points to the correct element")]
    public class MsCrmGetValueCommand : Command
    {
        public MsCrmGetValueCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Phrase to find the Element with", Required = true)]
            public TextStructure Search { get; set; }

            [Argument(Tooltip = "Specify the Selector: name, id, title, class name, selector, role of the element")]
            public TextStructure By { get; set; } = new TextStructure("id");

            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(MsCrmManager.Instance.CurrentCRM.GetValue(arguments.Search.Value, arguments.By.Value)));
        }

    }
}
