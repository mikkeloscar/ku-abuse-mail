package dk.ku.mail;

import java.io.Console;

public class Main {
    public static void main(String[] args) {
        try {
            Console con = System.console();

            if (con != null) {
                String username = con.readLine("Enter KU-username: ");
                char[] password = con.readPassword("Enter password: ");

                HandleAbuse abuse = new HandleAbuse(username + "@ku.dk", new String(password), "abuse@adm.ku.dk");
                abuse.connect();
                abuse.handleDuplicates();
                abuse.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
