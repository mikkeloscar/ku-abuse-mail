package dk.ku.mail;

import java.util.Scanner;

public class Main {
	public static void main(String [] args) {

        Scanner reader = new Scanner(System.in);
        System.out.print("Enter KU-username: ");
        String username = reader.nextLine();
        System.out.print("Enter password: ");
        String password = reader.nextLine();

        HandleAbuse abuse = new HandleAbuse(username + "@ku.dk", password, "abuse@adm.ku.dk");
		try {
            abuse.connect();
            abuse.handleDuplicates();
            abuse.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
