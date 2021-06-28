// io
import java.io.*;
import java.util.Properties;
import java.io.FileReader;
import java.io.IOException;

// dialog
import javax.swing.*;

// java mail
import javax.mail.*;
import javax.mail.internet.*;
import javax.swing.filechooser.FileNameExtensionFilter;

// opencsv
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;

public class SendEmail {
    static String htmlFilePath;
    static String csvFilePath;
    static StringBuilder htmlContent = new StringBuilder();
    static String subject;

    // authentication info
    final static String UserName = "a1155142308@yahoo.com";
    final static String Password = "ztnoupjpoegmaqfq";
    final static String FromEmail = "a1155142308@yahoo.com";

    public static void getFilePath(String fileType) {
        String filePath;
        String description;

        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Select your file");

        if (fileType.equals("html")) {
            description = "Select HTML file as email content";
        } else {
            description = "Select CSV file as email distribution list";
        }
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                description, fileType);
        fc.setFileFilter(filter);

        int returnVal = fc.showOpenDialog(null);
        if (returnVal == JFileChooser.APPROVE_OPTION && fileType.equals("html")) {
            File selectedFile = fc.getSelectedFile();
            System.out.println("html " + selectedFile.getAbsolutePath());
            filePath = selectedFile.getAbsolutePath();
            htmlFilePath = filePath;
        } else if (returnVal == JFileChooser.APPROVE_OPTION && fileType.equals("csv")) {
            File selectedFile = fc.getSelectedFile();
            System.out.println("csv " + selectedFile.getAbsolutePath());
            filePath = selectedFile.getAbsolutePath();
            csvFilePath = filePath;
        } else {
            System.exit(0);
        }
    }

    public static void sendEmail(MimeMessage msg, String subject) throws MessagingException, IOException {
        Address[] recipients = msg.getRecipients(Message.RecipientType.BCC);
        StringBuilder recipientStr = new StringBuilder();
        for (Address recipient : recipients) {
            recipientStr.append(recipient.toString()).append("\n");
        }
        String dialogMessage = String.format("HTML file path: %s\n"
                        + "CSV file path: %s\n"
                        + "Subject: %s\n"
                        + "Recipient: \n%s\n"
                        + "Are you sure to send the email?",
                htmlFilePath, csvFilePath, subject, recipientStr.toString());
        int option = JOptionPane.showConfirmDialog(null, dialogMessage);
        if (option == JOptionPane.YES_OPTION) {
            Transport.send(msg);
            JOptionPane.showMessageDialog(null, "Email sent successfully");
        } else {
            System.exit(0);
        }
    }

    public static void readRecipientsFromCsv(MimeMessage msg) throws IOException, CsvValidationException, MessagingException {
        CSVReader reader = new CSVReader(new FileReader(csvFilePath));
        String[] nextLine;
        boolean flag = false;
        while ((nextLine = reader.readNext()) != null) {
            if (!nextLine[0].contains("@")) {
                continue;
            }
            System.out.println(nextLine[0]);
            msg.addRecipient(Message.RecipientType.BCC, new InternetAddress(nextLine[0]));
            flag = true;
        }
        if (!flag) {
            JOptionPane.showMessageDialog(null, "No email address at column 1 of CSV file");
            System.exit(0);
        }
    }

    public static Session getSession(Properties properties) {
        return Session.getInstance(properties, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(UserName, Password);
            }
        });
    }

    public static void readHtmlFile(StringBuilder htmlContent) {
        try {
            BufferedReader br = new BufferedReader(new FileReader(htmlFilePath));
            String thisLine;
            while ((thisLine = br.readLine()) != null) {
                htmlContent.append(thisLine);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Properties getProperties() {
        Properties properties = new Properties();
        properties.put("mail.mime.address.usecanonicalhostname", "false");
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.starttls.enable", "true");
        properties.put("mail.smtp.host", "smtp.mail.yahoo.com");
        properties.put("mail.smtp.port", "587");
        return properties;
    }

    public static void setSubject(MimeMessage msg) throws MessagingException {
        subject = JOptionPane.showInputDialog("Please input email subject");
        if (subject == null) {
            System.exit(0);
        }
        msg.setSubject(subject);
    }

    /*
    main program
     */
    public static void main(String[] args) {
        // read html file
        getFilePath("html");
        readHtmlFile(htmlContent);

        // set email content
        Properties properties = getProperties();
        Session session = getSession(properties);
        MimeMessage msg = new MimeMessage(session);
        try {
            msg.setFrom(new InternetAddress(FromEmail));
            // read recipients
            getFilePath("csv");
            readRecipientsFromCsv(msg);

            // set email subject
            setSubject(msg);

            // set html content
            msg.setContent(htmlContent.toString(), "text/html");

            // send email
            sendEmail(msg, subject);
        } catch (MessagingException | CsvValidationException | IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Sending failed", "Alert", JOptionPane.WARNING_MESSAGE);
            System.exit(0);
        }
    }

}