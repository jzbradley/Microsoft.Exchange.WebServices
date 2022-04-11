using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;
using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices
{
    public static class SearchExtensions
    {
        public static IEnumerable<Folder> QueryFolders(this ExchangeService service, Action<FolderQuery> query)
        {
            var setup = new FolderQuery(service);
            query(setup);
            return setup;
        }
        public static IEnumerable<T> Query<T>(this ExchangeService service, Action<ItemQuery<T>> query)
        where T:Item
        {
            var setup = new ItemQuery<T>(service);
            query(setup);
            return setup;
        }

        public static IEnumerable<EmailMessage> QueryMessages(this ExchangeService service,
            Action<ItemQuery<EmailMessage>> query)
            => service.Query(query);
    }
}