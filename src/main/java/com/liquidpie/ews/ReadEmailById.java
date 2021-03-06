package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class ReadEmailById {

	public static void main(String[] args) throws Exception {

		ExchangeService service = ConnectionUtil.getService();
		FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
				new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(100));  
		for(Item item: results){
			EmailMessage message = EmailMessage.bind(service, item.getId());
			System.out.println(message.getBody());
		}
	}
}
