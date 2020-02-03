using System;
using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Command(Name = "mscrm.click",Tooltip = "This command allows to send click event to element of an active CRM instance.")]
    public class MsCrmClickCommand : Command
    {
        public MsCrmClickCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Phrase to find the Element with", Required = true)]
            public TextStructure Search { get; set; }

            [Argument(Tooltip = "Specify the Selector: name, id, title, class name, selector, role of the element")]
            public TextStructure By { get; set; } = new TextStructure("id");

            [Argument]
            public BooleanStructure Trigger { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Whether to wait for the element to appear on screen")]
            public BooleanStructure NoWait { get; set; } = new BooleanStructure(false);

            [Argument(DefaultVariable = "timeoutcrm")]
            public override TimeSpanStructure Timeout { get; set; }             
        }

        public void Execute(Arguments arguments)
        {
            if (MsCrmManager.Instance.CurrentCRM != null)
            {
                MsCrmManager.Instance.CurrentCRM.ClickByElement(arguments.Search.Value, arguments.By.Value, arguments.Trigger.Value, arguments.NoWait.Value);
                if (!arguments.NoWait.Value)
                {
                    MsCrmManager.Instance.CurrentCRM.Ie.WaitForComplete();
                }
            }
            else
            {
                throw new ApplicationException("Unable to find CRM. Use mscrm.attach first.");
            }
        }
    }
}
