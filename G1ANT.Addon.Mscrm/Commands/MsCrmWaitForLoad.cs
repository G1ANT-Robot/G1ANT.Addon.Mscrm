using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.waitforload")]
    public class MsCrmWaitForLoadCommand : Command
    {
        public MsCrmWaitForLoadCommand(AbstractScripter scripter) : base(scripter) { }
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Phrase to find the Element with", Required = true)]
            public TextStructure Search { get; set; }

            [Argument(Tooltip = "Specify the Selector: name, id, title, class name, selector, role of the element")]
            public TextStructure By { get; set; } = new TextStructure("id");

            [Argument]
            public override TimeSpanStructure Timeout { get; set; } = new TimeSpanStructure(10000);
        }

        public void Execute(Arguments arguments)
        {
            var currentCrm = MsCrmManager.Instance.CurrentCRM;
            currentCrm.WaitForLoad(arguments.Search.Value, arguments.By.Value, (int)arguments.Timeout.Value.TotalMilliseconds);            
        }
    }
}
