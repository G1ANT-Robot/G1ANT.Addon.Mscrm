using G1ANT.Language;

namespace G1ANT.Addon.Mscrm
{
    [Addon(Name = "mscrm", Tooltip = "MsCrm Commands")]
    [Copyright(Author = "G1ANT LTD", Copyright = "G1ANT LTD", Email = "hi@g1ant.com", Website = "www.g1ant.com")]
    [License(Type = "LGPL", ResourceName = "License.txt")]
    [CommandGroup(Name = "mscrm", Tooltip = "Command connected with creating, editing and generally working on mscrm")]
    public class MsCrmAddon : Language.Addon { }
}