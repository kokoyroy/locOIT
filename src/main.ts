/// <reference types="@types/office-js" />

import { bootstrapApplication } from '@angular/platform-browser';
import { AppComponent } from './app/app';
import { appConfig } from './app/app.config';

declare var Office: any;

function startApp() {
  console.log('🚀 Starting Angular application...');
  bootstrapApplication(AppComponent, appConfig)
    .then(() => console.log('✅ Angular application started successfully'))
    .catch((err) => {
      console.error('❌ Error starting Angular application:', err);
    });
}

// Check if Office.js is available
if (typeof Office !== 'undefined') {
  console.log('🏢 Office.js detected');

  // Use Office.onReady (modern approach)
  Office.onReady((info: any) => {
    console.log('Office.onReady triggered:', info);

    if (info.host === Office.HostType.Outlook) {
      console.log('🏢 Office.js initialized for Outlook');
    } else {
      console.log('🏢 Office.js initialized for other host:', info.host);
    }

    startApp();
  });

  // Fallback to legacy initialization if onReady doesn't work
  setTimeout(() => {
    if (!Office.context) {
      console.log('🏢 Using legacy Office.initialize fallback');
      Office.initialize = () => {
        console.log('🏢 Office.js initialized (legacy)');
        startApp();
      };
    }
  }, 1000);
} else {
  // Fallback for testing outside of Office environment
  console.log('⚠️ Office.js not available - running in standalone mode');
  startApp();
}
