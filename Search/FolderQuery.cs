using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices
{
    public class FolderQuery : IEnumerable<Folder>
    {
        private readonly ExchangeService _service;
        public FolderView View { get; set; } = new FolderView(100);
        public Mailbox Mailbox { get; set; } = new Mailbox();
        public SearchFilter SearchFilter { get; set; }
        public List<FolderId> Folders { get; } = new List<FolderId>();

        public FolderQuery(ExchangeService service)
        {
            _service = service;
        }
        public IEnumerator<Folder> GetEnumerator()
        {
            var folders
                = Folders.Count == 0 ? new[] { new FolderId(WellKnownFolderName.Root, Mailbox) }
                : string.IsNullOrEmpty(Mailbox?.Address) ? Folders.ToArray()
                : Folders.Select(folder => new FolderId(folder.FolderName ?? WellKnownFolderName.Root, Mailbox))
                    .ToArray();
            var view = new FolderView(View.PageSize, 0, View.OffsetBasePoint)
            {
                PropertySet = View.PropertySet,
                Traversal = View.Traversal
            };

            FindFoldersResults results;
            do
            {
                results = _service.FindFolders(folders, SearchFilter, View)[0].Results;

                foreach (var item in results)
                    yield return item;
                view.Offset += view.PageSize;

            } while (results.MoreAvailable);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}