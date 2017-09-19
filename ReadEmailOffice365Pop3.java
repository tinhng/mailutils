package tinhn.org;

/**
 * Created by tinhng on 9/14/2017.
 */

import com.google.gson.Gson;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Properties;


import javax.mail.*;
import javax.mail.internet.MimeMultipart;
import javax.mail.search.FlagTerm;

public class ReadEmailOffice365Pop3 {

    private static Logger _logger = LoggerFactory.getLogger(ReadEmailOffice365Pop3.class);    

    public static void ReadMail(String host, String storeType, String user, String password) {       
        try {

            //create properties field
            Properties properties = new Properties();

            properties.put("mail.pop3.host", host);
            properties.put("mail.pop3.port", "995");
            //properties.put("mail.pop3.starttls.enable", "true");
            //properties.setProperty("mail.pop3.socketFactory.fallback", "false");
            //properties.setProperty("mail.pop3.socketFactory.port",String.valueOf("995"));
            //properties.put("mail.pop3.auth", "false");
            //properties.put("mail.debug.auth", "true");
            properties.put("mail.pop3s.ssl.trust", "*"); // This is the most IMP property

            Session emailSession = Session.getDefaultInstance(properties);

            //create the POP3 store object and connect with the pop server

            Store store = emailSession.getStore("pop3s"); //try imap or impas
            store.connect(host, user, password);

            //create the folder object and open it
            Folder emailFolder = store.getFolder("INBOX");
            emailFolder.open(Folder.READ_ONLY);

            // retrieve the messages from the folder in an array and print it
            Message[] messages = emailFolder.getMessages();
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
        String mailStoreType = "pop3s";
        String username = "username";// change accordingly
        String password = "password";// change accordingly

        ReadMail(host, mailStoreType, username, password);
    }    
}

