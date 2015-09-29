package dk.ku.mail;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import org.joda.time.DateTime;

import java.util.Date;

public class HandleAbuse {
    private String username;
    private String password;
    private String mailbox;
    private FolderId inbox;
    private FolderId handled;
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
        inbox = new FolderId(WellKnownFolderName.Inbox, mb);

        FindFoldersResults findResults = service.findFolders(inbox, new FolderView(5));

        for (Folder folder : findResults.getFolders()) {
            if (folder.getDisplayName().equals("Oprettet")) {
                handled = folder.getId();
            }
        }
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
                new SearchFilter.SearchFilterCollection(LogicalOperator.Or,
                        new SearchFilter.ContainsSubstring(ItemSchema.Subject, "DKCERT Abuse Report"),
                        new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Distribuering af ophavsretsbeskyttet")
                )
        );

        do {
            findResults = service.findItems(inbox, sf, view);

            for (Item item : findResults.getItems()) {
                markDuplicatesRead(item.getSubject());
                count++;
            }

            view.setOffset(view.getOffset() + offset);
        } while (findResults.isMoreAvailable());

        System.out.println(count);
    }

    private void markDuplicatesRead(String subject) throws Exception {
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

        FindItemsResults<Item> findResults = service.findItems(inbox, sf, view);

        EmailMessage msg;
        int i = -1;
        for (Item item : findResults.getItems()) {
            i++;
            msg = (EmailMessage) item;

            if (i == 0) {
                // Check if report was handled within the last 14 days.
                markAlreadyHandled(sub, msg);
                continue;
            }

            msg.setIsRead(true);
            msg.update(ConflictResolutionMode.AutoResolve);
            System.out.println("Mark report read: " + msg.getSubject());
        }
    }

    private void markAlreadyHandled(String subject, EmailMessage msg) throws Exception {
        ItemView view = new ItemView(1);
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        view.setPropertySet(
                new PropertySet(
                        BasePropertySet.IdOnly,
                        ItemSchema.Subject,
                        ItemSchema.DateTimeReceived)
        );

        SearchFilter sf = new SearchFilter.SearchFilterCollection(
                LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true),
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, subject)
        );

        FindItemsResults<Item> findResults = service.findItems(handled, sf, view);

        // 14 days ago
        Date after = new DateTime(new Date()).minusDays(14).toDate();

        if (findResults.getItems().size() == 1 &&
                findResults.getItems().get(0).getDateTimeReceived().after(after)) {
            msg.setIsRead(true);
            msg.update(ConflictResolutionMode.AutoResolve);
            System.out.println("Mark report handled+read: " + msg.getSubject());
        }
    }
}