using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.setvalue", Tooltip = "This command allows to set value of a field in CRM form. This command allows to handle fields like: text, lookup, tree lookup and dropdown.")]
    public class MsCrmSetValueCommand : Command
    {
        public MsCrmSetValueCommand(AbstractScripter scripter) : base(scripter) { }
        public class Arguments : CommandArguments
        {
 
            [Argument(Tooltip = "Phrase to find the Element with", Required = true)]
            public TextStructure Search { get; set; }

            [Argument(Tooltip = "Specify the Selector: name, id, title, class name, selector, jquery, etc., of the element")]
            public TextStructure By { get; set; } = new TextStructure("id");

            [Argument(Required = true, Tooltip = "Enter the value to be set in the field")]
            public TextStructure Value { get; set; }

            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }
        }

        public void Execute(Arguments arguments)
        {
            var currentCrm = MsCrmManager.Instance.CurrentCRM;
            currentCrm.SetVal(arguments.Search.Value, arguments.Value.Value, arguments.By.Value);
        }
    }
}
