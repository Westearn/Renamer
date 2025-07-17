using System.Collections.Generic;


namespace Renamer
{
    public interface IRename
    {
        void DoRename(string directory, List<(string searchName, string replaceName)> renameList);
    }
}
