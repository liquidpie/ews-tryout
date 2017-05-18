package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.File;
import java.util.UUID;

public class ReadMailAttachment {
	public static void main(String[] args)  {
		try{
			ExchangeService service = ConnectionUtil.getService();

			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(1));  
			service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));

			for(Item item: results){
				EmailMessage msg = (EmailMessage)item;
				for(Attachment attachment:msg.getAttachments()){
					FileAttachment fileAttachment =(FileAttachment)attachment;
					File file= new File(UUID.randomUUID().toString() + attachment.getName());
					fileAttachment.load(file.getAbsolutePath());
				}
			}
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}
}
