package com.mvc;

import java.io.*;
import java.util.*;
import jakarta.activation.*;
import jakarta.mail.*;
import jakarta.mail.internet.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Mailer {
//LES FICHIERS DE CV , EMAILVALIDES ET progress.txt que j'ai cree ici ne sont pas utuilis
//depuis ce dossier mais depuis leurs chemins absolua
//il faut eviter d'envoyer un grand nombre des emails par une seue excution donc
// on fait une seuile de 15 emails a ne pas depasser par excution puus une paussse puis reprise
//puis aussi un nombre des emails qu'il faut pas depasser par jour ce nombre des emials
// se stockent aussi dans progress.txt comme le nombre des entreprise auquelles on a envoye
//puis pendant la journe auivante on fait regler le champ consacre au nombre des meails
// (je l'ai fixe qu'il faut pas depasser 50) sur 0 pour recommencer a nouveau(changer a la main)
    private static final String EXCEL_FILE    = "C:\\Users\\pc\\Downloads\\EMAILVALIDE.xlsx";
    private static final String CV_FILE       = "C:\\Users\\pc\\Downloads\\AOURIR-Aissam-CV-FR.pdf";
    private static final String PROGRESS_FILE = "C:\\Users\\pc\\Downloads\\progress.txt"; // 🔧 Chemin absolu ici

    private static final String SMTP_HOST  = "smtp.gmail.com";
    private static final int    SMTP_PORT  = 587;
    private static final String USER_EMAIL = "aissamaourir2@gmail.com";
    private static final String APP_PASS   = "gepi lqut nihq danh";
    private static final int MAX_PER_RUN = 15;
    private static final int MAX_TOTAL   = 50;

    public static void main(String[] args) {
        int lastLineSent = 0;
        int totalEmailsSent = 0;
        File progress = new File(PROGRESS_FILE);
        if (progress.exists()) {
            try (BufferedReader br = new BufferedReader(new FileReader(progress))) {
                String line1 = br.readLine();
                String line2 = br.readLine();
                if (line1 != null) lastLineSent = Integer.parseInt(line1.trim());
                if (line2 != null) totalEmailsSent = Integer.parseInt(line2.trim());
            } catch (Exception e) {
                System.err.println("Erreur lecture  : " + e.getMessage());
            }
        }

        int emailsSentThisRun = 0;
        int lastProcessedLine = lastLineSent;

        try (FileInputStream fis = new FileInputStream(EXCEL_FILE);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            int rowCount = sheet.getLastRowNum();

            outer:
            for (int i = lastLineSent + 1; i <= rowCount; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cEnt = row.getCell(0);
                Cell cMail = row.getCell(1);
                if (cEnt == null || cMail == null) continue;

                String entreprise = cEnt.getStringCellValue().trim();
                String rawEmails  = cMail.getStringCellValue().trim();

                if (rawEmails.isEmpty()) continue;

                String[] emails = rawEmails.split("[,;]");
                List<String> validEmails = new ArrayList<>();

                for (String email : emails) {
                    String clean = email.trim();
                    try {
                        InternetAddress emailAddr = new InternetAddress(clean, true);
                        validEmails.add(clean);
                    } catch (AddressException e) {
                        System.err.println("Email invalide ignoré : " + clean);
                    }
                }

                if (validEmails.isEmpty()) {
                    System.out.printf("Aucun email valide pour %s (ligne %d)%n", entreprise, i + 1);
                    continue;
                }

                for (String email : validEmails) {
                    if (totalEmailsSent >= MAX_TOTAL) {
                        System.out.println("⛔ Limite totale atteinte (" + MAX_TOTAL + ").");
                        break outer;
                    }
                    if (emailsSentThisRun >= MAX_PER_RUN) {
                        System.out.println("⏳ Limite par exécution atteinte (" + MAX_PER_RUN + ").");
                        break outer;
                    }

                    try {
                        sendMail(entreprise, email);
                        emailsSentThisRun++;
                        totalEmailsSent++;
                        System.out.printf("✓ Envoyé à %s <%s> (ligne %d)%n", entreprise, email, i + 1);
                    } catch (MessagingException me) {
                        System.err.printf("✗ Échec pour %s <%s> : %s%n", entreprise, email, me.getMessage());
                    }

                    try {
                        Thread.sleep(1500); // Pause entre les envois
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                    }
                }

                lastProcessedLine = i;
            }

        } catch (Exception e) {
            System.err.println("Erreur fatale : " + e.getMessage());
            return;
        }

        // Sauvegarde finale de la progression
        try (PrintWriter pw = new PrintWriter(new FileWriter(PROGRESS_FILE))) {
            pw.println(lastProcessedLine);
            pw.println(totalEmailsSent);
            System.out.println("📄 Progression mise à jour : ligne " + lastProcessedLine + ", total " + totalEmailsSent);
        } catch (IOException e) {
            System.err.println("❌ Erreur écriture progress.txt : " + e.getMessage());
        }

        System.out.printf("✅ Fin de l'envoi : %d emails envoyés dans cette session. Total général : %d%n",
                emailsSentThisRun, totalEmailsSent);
    }

    private static void sendMail(String entreprise, String to) throws MessagingException {
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", SMTP_HOST);
        props.put("mail.smtp.port", String.valueOf(SMTP_PORT));

        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(USER_EMAIL, APP_PASS);
            }
        });

        MimeMessage msg = new MimeMessage(session);
        msg.setFrom(new InternetAddress(USER_EMAIL));
        msg.setRecipient(Message.RecipientType.TO, new InternetAddress(to));
        msg.setSubject("Demande de stage d'été en développement informatique.");

        MimeBodyPart textPart = new MimeBodyPart();
        String body = "Bonjour Madame, Monsieur,\n\n" +
                "Je suis étudiant à l'ENSA Marrakech en génie informatique, actuellement à la recherche d’un stage d'été.\n\n" +
                "Je suis très intéressé par un stage chez " + entreprise + " et vous transmets ci-joint mon CV.\n\n" +
                "Bien cordialement,\nAourir Aïssam\nTél. : 06 13 93 52 38\nEmail : " + USER_EMAIL;
        textPart.setText(body, "UTF-8");

        MimeBodyPart filePart = new MimeBodyPart();
        DataSource source = new FileDataSource(new File(CV_FILE));
        filePart.setDataHandler(new DataHandler(source));
        filePart.setFileName(new File(CV_FILE).getName());

        Multipart mp = new MimeMultipart();
        mp.addBodyPart(textPart);
        mp.addBodyPart(filePart);
        msg.setContent(mp);

        Transport.send(msg);
    }
}
