package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class ForwardMail {

	public static void main(String[] args) throws Exception {
		ExchangeService service = ConnectionUtil.getService();

		FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox,
				new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(100));  
		for(Item item: results){
			EmailMessage msg = EmailMessage.bind(service, item.getId());
			MessageBody messageBody = new MessageBody("");
			EmailAddress emailAddress = new EmailAddress("user@example.com");
			msg.forward(messageBody, emailAddress);
		}
	}
	
}
