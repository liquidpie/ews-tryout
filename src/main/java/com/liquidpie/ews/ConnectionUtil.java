package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;

import java.net.URI;

public class ConnectionUtil {

	public static ExchangeService getService() throws Exception {
		ExchangeService service = new ExchangeService();
		ExchangeCredentials credentials = new WebCredentials("username",
				"password");
		service.setCredentials(credentials);
		service.setUrl(new URI("https://Webmail.companyname.com/ews/Exchange.asmx"));

		return service;
	}
}
