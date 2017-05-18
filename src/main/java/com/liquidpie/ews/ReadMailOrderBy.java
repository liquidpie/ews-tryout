package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class ReadMailOrderBy {

	public static void main(String[] args)  {
		try{
			ExchangeService service = ConnectionUtil.getService();
			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(100));  
			service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));

			for(Item item: results){
				EmailMessage msg = (EmailMessage)item;
				System.out.println(msg.getSubject());
			}
			System.out.println("Reading mailbox in the mail incoming order...");

			ItemView itemView  = new ItemView(5);
			itemView.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending); // Sort mails in the incoming order
			FindItemsResults<Item> results1 = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), itemView);
			service.loadPropertiesForItems(results1, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));

			for(Item item: results1){
				EmailMessage msg = (EmailMessage)item;
				System.out.println(msg.getSubject());
			}

		}
		catch(Exception e){
			e.printStackTrace();
		}
	}

}
