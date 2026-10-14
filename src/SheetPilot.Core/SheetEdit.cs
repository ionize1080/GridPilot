namespace SheetPilot.Core
{
    public sealed class SheetEdit
    {
        public SheetItem Source {get;set;}
        public SheetItem Target {get;set;}
        public string Action {get;set;}
        public string Name {get;set;}
        public bool After {get;set;}
    }
}
