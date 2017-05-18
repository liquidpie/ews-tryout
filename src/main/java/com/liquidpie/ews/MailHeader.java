package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.DefaultExtendedPropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.property.complex.ExtendedProperty;
import microsoft.exchange.webservices.data.property.complex.InternetMessageHeader;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class MailHeader {
	
	public static void main(String[] args) throws Exception{
		try{
			ExchangeService service = ConnectionUtil.getService();

			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(1));  
			service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));

			for(Item item: results){
				EmailMessage msg = (EmailMessage)item;
				
				for(InternetMessageHeader header: msg.getInternetMessageHeaders().getItems()){
					System.out.println(header.getName()+": "+header.getValue());
					header.setValue("updated header value"); // update header value
					System.out.println(header.getName()+": "+header.getValue());
				}
				ExtendedPropertyDefinition headerElement = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "x-mycustom-header", MapiPropertyType.String);
			    msg.setExtendedProperty(headerElement, "My custom header value"); // add a custom header

			    for(ExtendedProperty header: msg.getExtendedProperties().getItems()){
			    	System.out.println(header.getValue());
			    }
			}
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}

}
