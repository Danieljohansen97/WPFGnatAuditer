namespace WPFGnatAuditer
{
    public class CiEntry
    {
        public int CiEntriesId { get; }
        public string Location { get; }
        public string SpecificLocation { get; }
        public string SubZone { get; }
        public string Site { get; }
        public string SubSite { get; }
        public string Component { get; }
        public string SubComponent { get; }
        public string Node { get; }
        public string ProbeSc { get; }
        public int CiPriority { get; }
        public int Type { get; }
        public int State { get; }
        public string CiDescription { get; }
        public string CiName { get; }

        public CiEntry(
            int ciEntriesId, string location, string specificLocation, string subZone,
            string site, string subSite, string component, string subComponent,
            string node, string probeSc, int ciPriority, int type, int state,
            string ciDescription, string ciName)
        {
            CiEntriesId = ciEntriesId;
            Location = location;
            SpecificLocation = specificLocation;
            SubZone = subZone;
            Site = site;
            SubSite = subSite;
            Component = component;
            SubComponent = subComponent;
            Node = node;
            ProbeSc = probeSc;
            CiPriority = ciPriority;
            Type = type;
            State = state;
            CiDescription = ciDescription;
            CiName = ciName;
        }
    }
}
