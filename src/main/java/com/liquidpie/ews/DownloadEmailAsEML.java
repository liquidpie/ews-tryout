package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class DownloadEmailAsEML {
	public static void main(String[] args)  {
		OutputStream outputStream = null;
		try {
			ExchangeService service = ConnectionUtil.getService();

			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(1));  
			service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
			
			System.out.println(results.getTotalCount());

			Item item = results.getItems().get(0); // writing only first email
			EmailMessage msg = (EmailMessage)item;
			msg.load(new PropertySet(ItemSchema.MimeContent));
			byte[] buffer = msg.getMimeContent().getContent();
			String fileName = "filename.eml";
			outputStream = new FileOutputStream(fileName);
			outputStream.write(buffer);
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally {
			try {
				outputStream.close();
			}
			catch (IOException e) {
				// Close quietly
			}
		}
	}
}
