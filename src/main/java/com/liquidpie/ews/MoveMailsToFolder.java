package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.util.ArrayList;
import java.util.List;

public class MoveMailsToFolder {

	public static void main(String[] args)  {
		try{
			ExchangeService service = ConnectionUtil.getService();

			Folder root = Folder.bind(service, WellKnownFolderName.MsgFolderRoot);
			FindFoldersResults folders = service.findFolders(root.getId(), new FolderView(20));
			folders.getFolders();
			Folder destFolder = null;
			for(Folder folder: folders.getFolders()){
				System.out.println(folder.getDisplayName());
				if(folder.getDisplayName().equalsIgnoreCase("destfolder-name")){
					destFolder = folder;
				}
			}

			FindItemsResults<Item> results = service.findItems(WellKnownFolderName.Inbox, 
					new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, "false"), new ItemView(100));  
			service.loadPropertiesForItems(results, PropertySet.FirstClassProperties);

			// Move mails to destFolder
			List<ItemId> itemids = new ArrayList<ItemId>();
			for(Item item: results){
				itemids.add(item.getId());
			}
			service.moveItems(itemids, destFolder.getId());

		}
		catch(Exception e){
			e.printStackTrace();
		}
	}

}
