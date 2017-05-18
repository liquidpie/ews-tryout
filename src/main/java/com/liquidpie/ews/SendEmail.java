package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.DefaultExtendedPropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;

public class SendEmail {

	public static void main(String[] args) throws Exception {
		ExchangeService service = ConnectionUtil.getService();
		
		try {
			
			EmailMessage msg = new EmailMessage(service);
			EmailAddress fromEmailAddress = new EmailAddress("user@example.com");
			msg.setFrom(fromEmailAddress);
			msg.getToRecipients().add("user2@example.com");
			msg.setSubject("Custom Header Test Mail");
			msg.setBody(new MessageBody("Test Mail"));
			
			// Add custom header
			ExtendedPropertyDefinition headerElement = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "x-mycustom-header", MapiPropertyType.String);
		    msg.setExtendedProperty(headerElement, "My custom header value");
			
			msg.send();
			
		} catch (Exception e) {
			e.printStackTrace();
		}		
	}
}
