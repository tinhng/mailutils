package tinhn.org;

/**
 * Created by tinhng
 */

import java.util.Properties;


import javax.mail.*;
import javax.mail.search.FlagTerm;

public class ConnectOutlook365Imap {
    public static String username = null;
    public static String password1 = null;

    public static void ReadEmail(String host, String storeType, String user,
                             String password) {
        username = user;
        password1 = password;
        try {
            //create properties field
            Properties properties = new Properties();

            properties.put("mail.imap.host", host);
            properties.put("mail.imap.port", "993");
            //properties.put("mail.imap.starttls.enable", "true");
            //properties.setProperty("mail.imap.socketFactory.fallback", "false");
            //properties.setProperty("mail.imao.socketFactory.port",String.valueOf("993"));
            //properties.put("mail.imap.auth", "false");
            //properties.put("mail.debug.auth", "true");
            properties.put("mail.imaps.ssl.trust", "*"); // This is the most IMP property
            properties.put("mail.imap.ssl.enable", "true");

            Session emailSession = Session.getDefaultInstance(properties);


            //create the POP3 store object and connect with the pop server

            Store store = emailSession.getStore("imaps"); //try imap or impas
            store.connect(host, user, password);
            //  store.connect(host, 993, user, password);

            //create the folder object and open it
            Folder emailFolder = store.getFolder("INBOX");
            emailFolder.open(Folder.READ_WRITE);

            // retrieve the messages from the folder in an array and print it
            //search for all "unseen" messages
            Flags seen = new Flags(Flags.Flag.SEEN);
            FlagTerm unseenFlagTerm = new FlagTerm(seen, false);
            Message messages[] = emailFolder.search(unseenFlagTerm);

            //Message[] messages = emailFolder.getMessages();
            System.out.println("messages.length---" + messages.length);

            for (int i = 0, n = messages.length; i < 10; i++) {
                Message message = messages[i];
                System.out.println("---------------------------------");
                System.out.println("Email Number " + (i + 1));
                System.out.println("Subject: " + message.getSubject());
                System.out.println("From: " + message.getFrom()[0]);
                System.out.println("Text: " + message.getContent().toString());

            }

            //close the store and folder objects
            emailFolder.close(false);
            store.close();

        } catch (NoSuchProviderException e) {
            e.printStackTrace();
        } catch (MessagingException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        String host = "outlook.office365.com";// change accordingly
        String mailStoreType = "imaps";
        String username = "username";// change accordingly
        String password = "password";// change accordingly

        ReadEmail(host, mailStoreType, username, password);

    }

}
