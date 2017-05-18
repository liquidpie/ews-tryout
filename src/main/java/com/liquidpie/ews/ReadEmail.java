package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.property.complex.*;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.File;
import java.util.UUID;

public class ReadEmail {
	public static void main(String[] args)  {
		try{
			ExchangeService service = ConnectionUtil.getService();

			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.From, "user@example.com"), new ItemView(1));
			
			service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));

			PropertySet propertySet = new PropertySet(BasePropertySet.FirstClassProperties);
			propertySet.setRequestedBodyType(BodyType.Text);
			propertySet.add(ItemSchema.MimeContent);

			for(Item item: results){
				item.load(propertySet);
				EmailMessage msg = (EmailMessage)item;
				System.out.println(msg.getMimeContent().toString());
				MessageBody body = msg.getBody();
				System.out.println(MessageBody.getStringFromMessageBody(body));

				// Get Headers
				for(InternetMessageHeader header: msg.getInternetMessageHeaders().getItems()){
					System.out.println(header.getName()+": "+ header.getValue());
				}
				
				for(ExtendedProperty header: msg.getExtendedProperties().getItems()){
			    	System.out.println(header.getValue());
			    }
				
				// Get From, To, CC, BCC
				System.out.println("From: " + msg.getFrom().getAddress());
				System.out.println("To: " + msg.getToRecipients().getItems().get(0));
				System.out.println("CC: " + msg.getCcRecipients());
				System.out.println("BCC: " + msg.getBccRecipients());
				
				// Get Subject
				System.out.println("Subject: " + msg.getSubject());
				
				// Get Body
				System.out.println("Body: " + msg.getBody());
				
				// Get Attachments
				for(Attachment attachment : msg.getAttachments()){
					FileAttachment fileAttachment =(FileAttachment) attachment;
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
