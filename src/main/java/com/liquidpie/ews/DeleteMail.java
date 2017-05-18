package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class DeleteMail {

	public static void main(String[] args)  {
		try{
			ExchangeService service = ConnectionUtil.getService();

			// Search for mail
			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(100));  
			service.loadPropertiesForItems(results, PropertySet.FirstClassProperties);
			System.out.println(results.getTotalCount()); // total number of mails found
			for(Item item: results){
				EmailMessage msg = (EmailMessage)item;
				/**
				 * there are modes of delete
				 * DeleteMode.MoveToDeletedItems
				 * DeleteMode.HardDelete
				 * DeleteMode.SoftDelete
				 */
				msg.delete(DeleteMode.MoveToDeletedItems);
			}
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}
}
