/**
 * Installation script for automatic analysis and artist breakdown updates
 * This script helps set up all necessary triggers
 */

/**
 * Run this function to install all automatic updating features
 */
function installAutomaticUpdates() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirm with user
  const response = ui.alert(
    'Install Automatic Updates',
    'This will set up triggers to automatically update your Analysis Dashboard and Artist Breakdown ' +
    'whenever a new "Attended" concert is added, and also at 2:00 AM daily.\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Installation Cancelled', 'No changes were made.', ui.ButtonSet.OK);
    return;
  }
  
  // Proceed with installation
  try {
    // Set up all triggers
    setupAllTriggers();
    
    // Ensure Artist Breakdown is set up
    ensureArtistBreakdownExists();
    
    // Run the analysis once to make sure everything is up to date
    updateAnalysis();
    
    // Show success message
    ui.alert(
      'Installation Successful',
      'Automatic updates have been set up successfully!\n\n' +
      'Your Concert Analysis and Artist Breakdown will now update:\n' +
      '1. Whenever a new "Attended" entry is added\n' +
      '2. Daily at 2:00 AM (server time)\n\n' +
      'You can still run these updates manually from the Concert Tracker menu.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    // Show error message
    ui.alert(
      'Installation Error',
      'An error occurred during installation: ' + error.toString() + '\n\n' +
      'Please try again or contact the script author for assistance.',
      ui.ButtonSet.OK
    );
    console.error('Installation error:', error);
  }
}

