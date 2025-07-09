import { Component, OnInit, NgZone } from '@angular/core';
import { CommonModule } from '@angular/common';
import { OutlookEventsService } from './services/outlook-events.service';

declare var Office: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.html',
  styleUrls: ['./app.css'],
  standalone: true,
  imports: [CommonModule],
})
export class AppComponent implements OnInit {
  title = 'Outlook Angular Add-in';
  isOfficeReady = false;
  currentEmail: string = 'No email selected';
  currentSubject: string = 'No subject';
  eventLog: string[] = [];
  debugInfo: string = '';

  constructor(
    private outlookEventsService: OutlookEventsService,
    private ngZone: NgZone,
  ) {}

  ngOnInit() {
    console.log('AppComponent initializing...');
    this.initializeOffice();
  }

  private initializeOffice() {
    // Check if Office is available
    if (typeof Office === 'undefined') {
      console.error('Office.js is not loaded');
      this.debugInfo =
        'Error: Office.js is not loaded. Please ensure you are running this add-in within Outlook.';
      return;
    }

    console.log('Office object detected, initializing...');
    this.debugInfo = 'Office.js detected, initializing...';

    try {
      Office.onReady((info: any) => {
        this.ngZone.run(() => {
          console.log('Office.onReady called', info);

          if (info.host === Office.HostType.Outlook) {
            console.log('Running in Outlook');
            this.isOfficeReady = true;
            this.debugInfo = `Office ready! Host: ${info.host}, Platform: ${info.platform}`;
            this.initializeOutlookFeatures();
          } else {
            console.warn('Not running in Outlook, but Office is ready');
            this.debugInfo = `Office ready, but not in Outlook. Host: ${info.host}`;
            this.isOfficeReady = true; // Still mark as ready for testing
          }
        });
      });
    } catch (error) {
      console.error('Error initializing Office:', error);
      this.debugInfo = `Error initializing Office: ${error}`;
    }

    // Fallback timeout
    setTimeout(() => {
      if (!this.isOfficeReady) {
        console.warn('Office initialization timeout, enabling fallback mode');
        this.ngZone.run(() => {
          this.debugInfo =
            'Office initialization timeout - running in fallback mode';
          this.isOfficeReady = true;
        });
      }
    }, 5000);
  }

  private initializeOutlookFeatures() {
    try {
      if (Office.context?.mailbox?.item) {
        console.log('Mailbox item available');
        this.loadCurrentEmail();
        this.setupEventHandlers();
      } else {
        console.log('No mailbox item available');
        this.debugInfo += ' | No email item selected';
      }
    } catch (error) {
      console.error('Error initializing Outlook features:', error);
      this.debugInfo += ` | Error: ${error}`;
    }
  }

  private loadCurrentEmail() {
    try {
      const item = Office.context.mailbox.item;
      if (item) {
        // Get sender email
        if (item.from) {
          this.currentEmail = item.from.emailAddress || 'Unknown sender';
        } else if (item.to && item.to.length > 0) {
          this.currentEmail = item.to[0].emailAddress || 'Unknown recipient';
        }

        // Get subject
        this.currentSubject = item.subject || 'No subject';

        this.addToEventLog(`Email loaded: ${this.currentEmail}`);
        console.log('Current email loaded:', this.currentEmail);
      }
    } catch (error) {
      console.error('Error loading current email:', error);
      this.addToEventLog(`Error loading email: ${error}`);
    }
  }

  private setupEventHandlers() {
    try {
      // Log that event handlers are being set up
      console.log('Event handlers set up successfully');
      this.addToEventLog('Event monitoring initialized');
    } catch (error) {
      console.error('Error setting up event handlers:', error);
      this.addToEventLog(`Error setting up events: ${error}`);
    }
  }

  // UI Action Methods
  showSubject() {
    if (!this.isOfficeReady) {
      alert('Office is not ready yet. Please wait...');
      return;
    }

    try {
      const item = Office.context?.mailbox?.item;
      if (item?.subject) {
        alert(`Subject: ${item.subject}`);
        this.addToEventLog(`Showed subject: ${item.subject}`);
      } else {
        alert('No email subject available');
        this.addToEventLog('No subject available');
      }
    } catch (error) {
      console.error('Error showing subject:', error);
      alert(`Error: ${error}`);
    }
  }

  checkPinningSupport() {
    if (!this.isOfficeReady) {
      alert('Office is not ready yet. Please wait...');
      return;
    }

    try {
      // Check Office version and capabilities
      const context = Office.context;
      let diagnostics = [];

      // Basic Office info
      if (context.host) {
        diagnostics.push(`Host: ${context.host}`);
      }
      if (context.platform) {
        diagnostics.push(`Platform: ${context.platform}`);
      }
      if (context.requirements && context.requirements.isSetSupported) {
        const mailboxSupported = context.requirements.isSetSupported(
          'Mailbox',
          '1.15',
        );
        diagnostics.push(`Mailbox 1.15 supported: ${mailboxSupported}`);
      }

      // Check if we're in a taskpane
      if (context.ui && context.ui.messageParent) {
        diagnostics.push('Running in task pane');
      }

      // Check Outlook version
      if (context.mailbox && context.mailbox.diagnostics) {
        const diag = context.mailbox.diagnostics;
        if (diag.version) {
          diagnostics.push(`Outlook version: ${diag.version}`);
        }
        if (diag.hostName) {
          diagnostics.push(`Host name: ${diag.hostName}`);
        }
        if (diag.hostVersion) {
          diagnostics.push(`Host version: ${diag.hostVersion}`);
        }
      }

      const message = `📋 Pinning Diagnostics:\n\n${diagnostics.join('\n')}\n\n🔧 Troubleshooting:\n• Desktop Outlook: Should support pinning\n• Outlook on Web: Limited pinning support\n• Mobile: May not support pinning\n• Minimum version 1.15 required`;

      alert(message);
      this.addToEventLog('Pinning diagnostics checked');
    } catch (error) {
      console.error('Error checking pinning support:', error);
      alert(`Error checking pinning support: ${error}`);
    }
  }

  addCategory() {
    if (!this.isOfficeReady) {
      alert('Office is not ready yet. Please wait...');
      return;
    }

    try {
      const item = Office.context?.mailbox?.item;
      if (item && item.categories) {
        const newCategory = {
          displayName: 'Angular Add-in',
          color: Office.MailboxEnums.CategoryColor.Preset1,
        };

        item.categories.addAsync([newCategory], (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            this.addToEventLog('Category added successfully');
            alert('Category "Angular Add-in" added!');
          } else {
            console.error('Error adding category:', result.error);
            this.addToEventLog(
              `Error adding category: ${result.error.message}`,
            );
            alert(`Error adding category: ${result.error.message}`);
          }
        });
      } else {
        alert('Categories not supported or no item available');
        this.addToEventLog('Categories not supported');
      }
    } catch (error) {
      console.error('Error in addCategory:', error);
      alert(`Error: ${error}`);
    }
  }

  refreshData() {
    this.addToEventLog('Refreshing data...');

    if (this.isOfficeReady) {
      this.loadCurrentEmail();
      this.addToEventLog('Data refreshed');
    } else {
      this.initializeOffice();
    }
  }

  clearEventLog() {
    this.eventLog = [];
  }

  private addToEventLog(message: string) {
    const timestamp = new Date().toLocaleTimeString();
    this.eventLog.unshift(`[${timestamp}] ${message}`);

    // Keep only last 20 entries
    if (this.eventLog.length > 20) {
      this.eventLog = this.eventLog.slice(0, 20);
    }
  }
}
