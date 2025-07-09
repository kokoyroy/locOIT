import { Injectable } from '@angular/core';
import { BehaviorSubject } from 'rxjs';

export interface EmailItem {
  id: string;
  subject: string;
  from: string;
  to: string;
  dateReceived: Date;
  conversationId?: string;
}

@Injectable({
  providedIn: 'root',
})
export class OutlookEventsService {
  private isInitialized = false;
  private emailCollection: EmailItem[] = [];
  private emailCollectionSubject = new BehaviorSubject<EmailItem[]>([]);

  // Observable for components to subscribe to email collection changes
  public emailCollection$ = this.emailCollectionSubject.asObservable();

  constructor() {
    // Delay initialization to ensure Office.js is loaded
    setTimeout(() => {
      this.initializeEventListeners();
    }, 1000);
  }

  private initializeEventListeners(): void {
    try {
      if (
        typeof Office !== 'undefined' &&
        Office.context &&
        Office.context.mailbox
      ) {
        console.log('🔧 Initializing Outlook event listeners...');

        // Item selection changed event
        try {
          Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            this.onItemChanged.bind(this),
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('✅ ItemChanged event listener added successfully');
                this.isInitialized = true;
              } else {
                console.error(
                  '❌ Failed to add ItemChanged event listener:',
                  result.error,
                );
              }
            },
          );

          // Get current selected item info
          this.logCurrentItemInfo();
        } catch (error) {
          console.error('❌ Error setting up event listeners:', error);
        }

        // Listen for Office context events
        this.setupOfficeContextEvents();
      } else {
        console.warn(
          '⚠️ Office.js context not available - running in standalone mode',
        );
        this.isInitialized = true;
      }
    } catch (error) {
      console.error('❌ Error initializing Outlook events:', error);
      this.isInitialized = true; // Mark as initialized to prevent loops
    }
  }

  private onItemChanged(eventArgs: any): void {
    try {
      console.log('📧 ITEM CHANGED EVENT:', eventArgs);
      console.log('📧 Event type:', eventArgs.type);
      console.log('📧 Event source:', eventArgs.source);

      // Get details about the new selected item
      this.logCurrentItemInfo();
    } catch (error) {
      console.error('❌ Error handling item changed event:', error);
    }
  }

  private logCurrentItemInfo(): void {
    try {
      if (typeof Office !== 'undefined' && Office.context?.mailbox?.item) {
        const item = Office.context.mailbox.item;
        console.log('📋 CURRENT ITEM INFO:');
        console.log('  - Item ID:', item.itemId || 'Not available');
        console.log('  - Item Type:', item.itemType || 'Not available');
        console.log('  - Subject:', item.subject || 'Not available');
        console.log(
          '  - Date Created:',
          item.dateTimeCreated || 'Not available',
        );
        console.log(
          '  - Date Modified:',
          item.dateTimeModified || 'Not available',
        );

        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
          const messageItem = item as any;
          console.log('  - From:', messageItem.from || 'Not available');
          console.log('  - To:', messageItem.to || 'Not available');
          console.log('  - CC:', messageItem.cc || 'Not available');
          console.log(
            '  - Conversation ID:',
            messageItem.conversationId || 'Not available',
          );
        }
      } else {
        console.log('📭 No item currently selected or Office.js not available');
      }
    } catch (error) {
      console.error('❌ Error logging current item info:', error);
    }
  }

  private setupOfficeContextEvents(): void {
    try {
      if (typeof Office !== 'undefined' && Office.context) {
        console.log('🏢 Office Context Info:');
        console.log('  - Host:', Office.context.host || 'Not available');
        console.log(
          '  - Platform:',
          Office.context.platform || 'Not available',
        );
        console.log(
          '  - Requirements:',
          Office.context.requirements || 'Not available',
        );

        if (Office.context.mailbox) {
          console.log('📮 Mailbox Info:');
          console.log(
            '  - User Email:',
            Office.context.mailbox.userProfile?.emailAddress || 'Not available',
          );
          console.log(
            '  - User Name:',
            Office.context.mailbox.userProfile?.displayName || 'Not available',
          );
          console.log(
            '  - Time Zone:',
            Office.context.mailbox.userProfile?.timeZone || 'Not available',
          );
        }
      }
    } catch (error) {
      console.error('❌ Error setting up Office context events:', error);
    }
  }

  // Method to manually trigger current item logging
  public logCurrentSelection(): void {
    console.log('🔍 MANUAL SELECTION CHECK:');
    this.logCurrentItemInfo();
  }

  // Method to get selected items (for multi-select scenarios)
  public getSelectedItems(): void {
    try {
      if (typeof Office !== 'undefined' && Office.context?.mailbox) {
        // For newer Office.js versions that support multi-select
        const mailbox = Office.context.mailbox as any;
        if (mailbox.getSelectedItemsAsync) {
          mailbox.getSelectedItemsAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log('📬 SELECTED ITEMS:', result.value);
              result.value.forEach((item: any, index: number) => {
                console.log(`  Item ${index + 1}:`, item);
              });
            } else {
              console.log('📝 Single item selection mode or API not available');
              this.logCurrentItemInfo();
            }
          });
        } else {
          console.log(
            '📝 Multi-select API not available, using single selection',
          );
          this.logCurrentItemInfo();
        }
      } else {
        console.warn('⚠️ Office.js mailbox not available');
      }
    } catch (error) {
      console.error('❌ Error getting selected items:', error);
      this.logCurrentItemInfo();
    }
  }

  // Method to add currently selected email to collection
  public addCurrentEmailToCollection(): Promise<boolean> {
    return new Promise((resolve) => {
      try {
        if (typeof Office !== 'undefined' && Office.context?.mailbox?.item) {
          const item = Office.context.mailbox.item;

          const emailItem: EmailItem = {
            id: item.itemId || `temp-${Date.now()}`,
            subject: item.subject || 'No Subject',
            from: item.from?.emailAddress || 'Unknown Sender',
            to: Array.isArray(item.to)
              ? item.to.map((t) => t.emailAddress).join(', ')
              : 'Unknown',
            dateReceived: item.dateTimeCreated || new Date(),
            conversationId: (item as any).conversationId,
          };

          // Check if email already exists in collection
          const exists = this.emailCollection.some(
            (email) => email.id === emailItem.id,
          );

          if (!exists) {
            this.emailCollection.push(emailItem);
            this.emailCollectionSubject.next([...this.emailCollection]);
            console.log('✅ Email added to collection:', emailItem.subject);
            resolve(true);
          } else {
            console.log('⚠️ Email already in collection');
            resolve(false);
          }
        } else {
          console.warn('⚠️ No email selected or Office.js not available');
          resolve(false);
        }
      } catch (error) {
        console.error('❌ Error adding email to collection:', error);
        resolve(false);
      }
    });
  }

  // Method to remove email from collection
  public removeEmailFromCollection(emailId: string): void {
    this.emailCollection = this.emailCollection.filter(
      (email) => email.id !== emailId,
    );
    this.emailCollectionSubject.next([...this.emailCollection]);
    console.log('🗑️ Email removed from collection');
  }

  // Method to clear entire collection
  public clearEmailCollection(): void {
    this.emailCollection = [];
    this.emailCollectionSubject.next([]);
    console.log('🧹 Email collection cleared');
  }

  // Method to get collection count
  public getCollectionCount(): number {
    return this.emailCollection.length;
  }

  // Method to get all emails in collection
  public getEmailCollection(): EmailItem[] {
    return [...this.emailCollection];
  }

  // Method to process all emails in collection
  public processEmailCollection(): void {
    console.log('⚙️ Processing email collection:', this.emailCollection);
    // Add your processing logic here
    // For example: analyze emails, export data, apply actions, etc.
  }

  // Method to check if service is initialized
  public isServiceInitialized(): boolean {
    return this.isInitialized;
  }

  // Method to retry initialization
  public retryInitialization(): void {
    console.log('🔄 Retrying initialization...');
    this.isInitialized = false;
    this.initializeEventListeners();
  }
}
