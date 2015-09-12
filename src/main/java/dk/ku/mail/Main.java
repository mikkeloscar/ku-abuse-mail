package dk.ku.mail;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;

import java.util.Scanner;

public class Main {
	public static void main(String [] args) {

        Scanner reader = new Scanner(System.in);
        System.out.print("Enter KU-username: ");
        String username = reader.nextLine();
        System.out.print("Enter password: ");
        String password = reader.nextLine();

		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
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
