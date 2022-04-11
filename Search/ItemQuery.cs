using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices
{
    public class ItemQuery<TItem> : IEnumerable<TItem>
        where TItem : Item
    {
        private readonly ExchangeService _service;

        public ItemQuery(ExchangeService service)
        {
            _service = service;
        }
        public List<FolderId> Folders { get; } = new List<FolderId>();
        public Mailbox Mailbox { get; set; } = new Mailbox();
        public SearchFilter SearchFilter { get; set; }
        public string QueryString { get; set; }
        public ItemView View { get; set; } = new ItemView(100);
        public Grouping Grouping { get; set; }
        public IEnumerator<TItem> GetEnumerator()
        {
            var folders
                = Folders.Count == 0 ? new[] { new FolderId(WellKnownFolderName.Root, Mailbox) }
                : string.IsNullOrEmpty(Mailbox?.Address) ? Folders.ToArray()
                : Folders.Select(folder => new FolderId(folder.FolderName ?? WellKnownFolderName.Root, Mailbox))
                    .ToArray();
            var view = new ItemView(View.PageSize, 0, View.OffsetBasePoint)
            {
                PropertySet = View.PropertySet,
                Traversal = View.Traversal
            };

            FindItemsResults<TItem> results;
            do
            {
                results = _service.FindItems<TItem>(folders, SearchFilter, QueryString, View, Grouping,
                    ServiceErrorHandling.ThrowOnError)[0].Results;

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