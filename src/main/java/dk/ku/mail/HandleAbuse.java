package dk.ku.mail;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class HandleAbuse {
    private String username;
    private String password;
    private String mailbox;
    private FolderId folder;
    private ExchangeService service;

    public HandleAbuse(String username, String password, String mailbox) {
        this.username = username;
        this.password = password;
        this.mailbox = mailbox;
    }

    public void connect() throws Exception {
        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials(username, password);
        service.setCredentials(credentials);
        service.autodiscoverUrl(username);

        // setup shared mailbox
        Mailbox mb = new Mailbox(mailbox);
        folder = new FolderId(WellKnownFolderName.Inbox, mb);
    }

    public void close() {
        service.close();
    }

    public void handleDuplicates() throws Exception {
        int offset = 10;

        ItemView view = new ItemView(offset);
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        view.setPropertySet(
                new PropertySet(
                        BasePropertySet.IdOnly,
                        ItemSchema.Subject,
                        ItemSchema.DateTimeReceived)
        );

        FindItemsResults<Item> findResults;

        int count = 0;

        SearchFilter sf = new SearchFilter.SearchFilterCollection(
                LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, "DKCERT Abuse Report"));

        do {
            findResults = service.findItems(folder, sf, view);

            for (Item item : findResults.getItems()) {
                markDuplicatesRead(item.getSubject());
                count++;
            }

            view.setOffset(view.getOffset() + offset);
        } while (findResults.isMoreAvailable());

        System.out.println(count);
    }

    public void markDuplicatesRead(String subject) throws Exception {
        ItemView view = new ItemView(30);
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        view.setPropertySet(
                new PropertySet(
                        BasePropertySet.IdOnly,
                        ItemSchema.Subject,
                        ItemSchema.DateTimeReceived)
        );

        // remove unique DK-CERT ID
        String sub = subject.replaceAll("\\[DK-CERT #\\d+\\] ", "");

        SearchFilter sf = new SearchFilter.SearchFilterCollection(
                LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, sub)
        );

        FindItemsResults<Item> findResults = service.findItems(folder, sf, view);

        EmailMessage msg;
        int i = -1;
        for (Item item : findResults.getItems()) {
            i++;
            if (i == 0) {
                continue;
            }

            msg = (EmailMessage)item;
            msg.setIsRead(true);
            msg.update(ConflictResolutionMode.AutoResolve);
            System.out.println("Mark report unread: " + msg.getSubject());
        }
    }
}
