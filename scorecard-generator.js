/**
 * Modified Redpoint Competition Scorecard Generator
 * Creates individual scorecards with category-specific climb assignments
 */

// Add custom menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Competition Tools')
    .addItem('Generate Scorecards', 'generateScorecards')
    .addToUi();
}

function generateScorecards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the summary sheet data
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find "Summary" sheet. Please create it first.');
    return;
  }
  
  const eventName = summarySheet.getRange('B1').getValue();
  const numClimbs = summarySheet.getRange('B2').getValue();
  const numAttempts = summarySheet.getRange('B3').getValue();
  const eventType = summarySheet.getRange('B4').getValue();
  const scoringNotes = summarySheet.getRange('B5').getValue();
  
  if (!eventName || !numClimbs || !numAttempts || !eventType) {
    SpreadsheetApp.getUi().alert('Error: Please fill in Event Name (B1), Number of Routes (B2), Number of Attempts (B3), and Event Type (B4) in the Summary sheet.');
    return;
  }
  
  // Validate event type
  const eventTypeLower = eventType.toString().toLowerCase().trim();
  if (eventTypeLower !== 'boulder' && eventTypeLower !== 'rope') {
    SpreadsheetApp.getUi().alert('Error: Event Type (B4) must be either "boulder" or "rope".');
    return;
  }
  
  // Get climb assignments
  const climbsSheet = ss.getSheetByName('Climbs');
  if (!climbsSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find "Climbs" sheet. Please create it first.');
    return;
  }
  
  const climbAssignments = parseClimbAssignments(climbsSheet);
  
  // Get or create Queue Cards sheet
  let queueSheet = ss.getSheetByName('Queue Cards');
  if (queueSheet) {
    ss.deleteSheet(queueSheet);
  }
  queueSheet = ss.insertSheet('Queue Cards');
  
  // Hide gridlines for cleaner appearance
  queueSheet.setHiddenGridlines(true);
  
  // Parse CSV data
  const csvSheet = ss.getSheetByName('Registration');
  if (!csvSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find "Registration" sheet. Please upload your CSV first.');
    return;
  }
  
  const climbers = parseClimberData(csvSheet);
  
  if (climbers.validClimbers.length === 0) {
    SpreadsheetApp.getUi().alert('No registered climbers found.');
    return;
  }
  
  // Generate scorecards
  const skippedClimbers = createRedpointScorecards(queueSheet, climbers.validClimbers, climbAssignments, eventName, numClimbs, numAttempts, scoringNotes, eventType);
  
  // Combine all skipped climbers (from parsing and from scorecard creation)
  const allSkipped = [...climbers.skippedClimbers, ...skippedClimbers];
  
  // Generate check-in lists
  createCheckInLists(ss, climbers.validClimbers);
  
  // Report results
  let message = `Success! Generated ${climbers.validClimbers.length - skippedClimbers.length} scorecards in the "Queue Cards" sheet and check-in lists for each session.`;
  
  if (allSkipped.length > 0) {
    message += `\n\nWARNING: ${allSkipped.length} climber(s) were skipped due to errors:\n`;
    allSkipped.forEach(skipped => {
      message += `\n- ${skipped.name} (${skipped.category || 'Unknown'}): ${skipped.reason}`;
    });
  }
  
  SpreadsheetApp.getUi().alert(message);
}

function parseClimbAssignments(sheet) {
  const data = sheet.getDataRange().getValues();
  const assignments = {};
  
  // Skip header row, start at row 1
  for (let i = 1; i < data.length; i++) {
    const category = data[i][0];
    const climbsStr = data[i][1];
    
    if (!category || !climbsStr) continue;
    
    // Parse comma-separated climb numbers
    const climbs = climbsStr.toString().split(',').map(c => c.trim());
    
    // Store climb numbers for this category
    const cleanCategory = category.toString().trim();
    assignments[cleanCategory] = climbs;
    Logger.log(`Climb assignment: ${cleanCategory} -> ${climbs.join(', ')}`);
  }
  
  Logger.log(`Total categories with climb assignments: ${Object.keys(assignments).length}`);
  return assignments;
}

function parseClimberData(sheet) {
  const data = sheet.getDataRange().getValues();
  const climbers = [];
  const skippedClimbers = [];
  
  if (data.length < 2) {
    Logger.log('Not enough rows in data');
    return { validClimbers: climbers, skippedClimbers: skippedClimbers };
  }
  
  // Check if data is already split or needs splitting
  let headerRow;
  if (typeof data[0][0] === 'string' && data[0][0].includes(';')) {
    headerRow = data[0][0].split(';');
    Logger.log('Header (semicolon split): ' + headerRow.slice(0, 10).join(' | '));
  } else {
    headerRow = data[0];
    Logger.log('Header (already split): ' + headerRow.slice(0, 10).join(' | '));
  }
  
  // Find column indices
  const firstnameIdx = headerRow.indexOf('firstname');
  const lastnameIdx = headerRow.indexOf('lastname');
  const bibIdx = headerRow.indexOf('bib');
  
  Logger.log(`Column indices - firstname: ${firstnameIdx}, lastname: ${lastnameIdx}, bib: ${bibIdx}`);
  
  // Find all boulder ticket columns
  const ticketColumns = [];
  headerRow.forEach((header, idx) => {
if (header && (header.toString().endsWith(' boulder ticket') || header.toString().endsWith(' lead ticket') || header.toString().endsWith(' rope ticket')) && !header.toString().includes('status')) {      
const category = header.toString().replace(/ (boulder|lead|rope) ticket$/, '');
  ticketColumns.push({ index: idx, category: category });
    }
  });
  
  Logger.log(`Found ${ticketColumns.length} ticket columns: ${ticketColumns.map(t => t.category).join(', ')}`);
  
  // Process each climber row
  for (let i = 1; i < data.length; i++) {
    let row;
    if (typeof data[i][0] === 'string' && data[i][0].includes(';')) {
      row = data[i][0].split(';');
    } else {
      row = data[i];
    }
    
    const firstname = row[firstnameIdx] || '';
    const lastname = row[lastnameIdx] || '';
    const bib = row[bibIdx] || '';
    
    if (!firstname && !lastname) continue;
    
    const climberName = `${firstname} ${lastname}`.trim();
    let foundRegistration = false;
    
    // Find which category they're registered for
    for (let ticket of ticketColumns) {
      const ticketValue = row[ticket.index] || '';
      
      if (ticketValue && ticketValue.toString().trim() !== '') {
        const sessionMatch = ticketValue.toString().match(/\d+/);
        const session = sessionMatch ? sessionMatch[0] : ticketValue.toString();
        
        if (!session || session.trim() === '') {
          skippedClimbers.push({
            name: climberName,
            category: ticket.category,
            reason: 'No session information found'
          });
          foundRegistration = true;
          break;
        }
        
        climbers.push({
          firstname: firstname,
          lastname: lastname,
          name: climberName,
          bib: bib,
          category: ticket.category,
          session: session
        });
        
        Logger.log(`Added climber: ${firstname} ${lastname}, Category: ${ticket.category}, Session: ${session}`);
        foundRegistration = true;
        break;
      }
    }
    
    // If climber has a name but no registration found
    if (!foundRegistration && climberName) {
      skippedClimbers.push({
        name: climberName,
        category: null,
        reason: 'No registration/ticket information found'
      });
    }
  }
  
  Logger.log(`Total climbers found: ${climbers.length}, Skipped: ${skippedClimbers.length}`);
  
  // Sort by session, then category (youngest to oldest, F then M then U), then last name
  climbers.sort((a, b) => {
    // First by session
    const sessionA = parseInt(a.session) || 0;
    const sessionB = parseInt(b.session) || 0;
    if (sessionA !== sessionB) return sessionA - sessionB;
    
    // Then by category
    const catCompare = compareCategoriesRedpoint(a.category, b.category);
    if (catCompare !== 0) return catCompare;
    
    // Finally by last name
    return a.lastname.localeCompare(b.lastname);
  });
  
  return { validClimbers: climbers, skippedClimbers: skippedClimbers };
}

function compareCategoriesRedpoint(catA, catB) {
  // Extract gender and age
  const getGenderAge = (cat) => {
    const gender = cat.charAt(0); // F, M, or U
    const age = parseInt(cat.substring(1)) || 0;
    return { gender, age };
  };
  
  const a = getGenderAge(catA);
  const b = getGenderAge(catB);
  
  // Gender order: F, then M, then U
  const genderOrder = { 'F': 0, 'M': 1, 'U': 2 };
  const genderCompare = (genderOrder[a.gender] || 3) - (genderOrder[b.gender] || 3);
  if (genderCompare !== 0) return genderCompare;
  
  // Within same gender, younger first (smaller age number first)
  return a.age - b.age;
}

function createRedpointScorecards(sheet, climbers, climbAssignments, eventName, numClimbs, numAttempts, scoringNotes, eventType) {
  const MAX_ATTEMPTS = 10; // Maximum possible attempts
  const CARD_WIDTH_COLS = 1 + MAX_ATTEMPTS; // Always 11 columns (1 climb + 10 attempts)
  const ROWS_PER_PAGE = 63;
  
  const skippedClimbers = []; // Track climbers that couldn't be processed
  
  // Determine climb label prefix based on event type
  const eventTypeLower = eventType.toString().toLowerCase().trim();
  const climbPrefix = eventTypeLower === 'rope' ? 'Climb' : 'Boulder';
  
  // Calculate total rows needed
  const totalRowsNeeded = (32 * climbers.length) + 100;
  
  // Ensure sheet has enough rows and columns
  const currentRows = sheet.getMaxRows();
  if (totalRowsNeeded > currentRows) {
    sheet.insertRowsAfter(currentRows, totalRowsNeeded - currentRows);
  }
  
  if (sheet.getMaxColumns() < CARD_WIDTH_COLS + 1) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), (CARD_WIDTH_COLS + 1) - sheet.getMaxColumns());
  }
  
  let currentRow = 1;
  const STANDARD_ROW_HEIGHT = 21;
  
  let successfulCards = 0; // Track cards that were actually created
  
  for (let i = 0; i < climbers.length; i++) {
    const climber = climbers[i];
    
    try {
      const cardStartRow = currentRow;
      const isOddCard = (successfulCards % 2 === 0); // Use successfulCards for odd/even tracking
      
      let rowsUsedCount = 0;
      
      // Get climbs for this category
      const categoryClimbs = climbAssignments[climber.category] || [];
      const actualNumClimbs = Math.min(categoryClimbs.length, parseInt(numClimbs));
      
      // Check if climber has climbs assigned
      if (categoryClimbs.length === 0) {
        skippedClimbers.push({
          name: climber.name,
          category: climber.category,
          reason: 'No climbs assigned to category'
        });
        continue; // Skip this climber
      }
      
      Logger.log(`Card ${successfulCards + 1}: ${climber.name}, Category: ${climber.category}, Climbs found: ${categoryClimbs.length}, Actual climbs to show: ${actualNumClimbs}`);
    
    // Header section
    sheet.getRange(currentRow, 1).setValue('Competition:').setFontWeight('bold');
    sheet.getRange(currentRow, 2, 1, CARD_WIDTH_COLS - 1).merge().setValue(eventName);
    currentRow++;
    rowsUsedCount++;
    
    // Name row - double height
    sheet.getRange(currentRow, 1).setValue('Name:').setFontWeight('bold').setFontSize(16);
    sheet.getRange(currentRow, 2, 1, CARD_WIDTH_COLS - 1).merge().setValue(climber.name).setFontSize(16);
    sheet.setRowHeight(currentRow, STANDARD_ROW_HEIGHT * 2);
    currentRow++;
    rowsUsedCount += 2;
    
    // Category, Session, and Bib row - double height
    sheet.getRange(currentRow, 1).setValue('Category:').setFontWeight('bold').setFontSize(16);
    sheet.getRange(currentRow, 2, 1, 2).merge().setValue(climber.category).setFontSize(16);
    
    sheet.getRange(currentRow, 4).setValue('Session:').setFontWeight('bold').setFontSize(16);
    sheet.getRange(currentRow, 5, 1, 2).merge().setValue(climber.session).setHorizontalAlignment('left').setFontSize(16);
    
    sheet.getRange(currentRow, 7).setValue('Bib #:').setFontWeight('bold').setFontSize(16);
    sheet.getRange(currentRow, 8, 1, CARD_WIDTH_COLS - 7).merge().setValue(climber.bib).setHorizontalAlignment('left').setFontSize(16);
    sheet.setRowHeight(currentRow, STANDARD_ROW_HEIGHT * 2);
    currentRow++;
    rowsUsedCount += 2;
    
    // Blank row
    currentRow++;
    rowsUsedCount++;
    
    // Create climb grid with double-height rows
    const gridStartRow = currentRow;
    
    for (let climbIdx = 0; climbIdx < actualNumClimbs; climbIdx++) {
      const climbNumber = categoryClimbs[climbIdx];
      
      // Climb number column - use appropriate prefix based on event type
      sheet.getRange(gridStartRow + climbIdx, 1)
        .setValue(`${climbPrefix}: ${climbNumber}`)
        .setFontWeight('bold')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true);
      
      // Attempt columns
      for (let attemptNum = 1; attemptNum <= numAttempts; attemptNum++) {
        const cell = sheet.getRange(gridStartRow + climbIdx, 1 + attemptNum);
        cell.setValue(attemptNum);
        cell.setFontColor('#cccccc');
        cell.setFontSize(10);
        cell.setVerticalAlignment('top');
        cell.setHorizontalAlignment('left');
        cell.setBorder(true, true, true, true, true, true);
      }
      
      // Set row to double height
      sheet.setRowHeight(gridStartRow + climbIdx, STANDARD_ROW_HEIGHT * 2);
    }
    
    currentRow = gridStartRow + actualNumClimbs;
    rowsUsedCount += (actualNumClimbs * 2); // Each climb row is double-height
    
    // Scoring notes
    if (scoringNotes) {
      currentRow++;
      rowsUsedCount++;
      
      sheet.getRange(currentRow, 1).setValue('Scoring:').setFontWeight('bold');
      sheet.getRange(currentRow, 2, 1, CARD_WIDTH_COLS - 1).merge().setValue(scoringNotes).setWrap(true);
      currentRow++;
      rowsUsedCount++;
    }
    
    // Padding and spacing
    // With narrower width, more rows fit per page vertically
    // Need to add more padding to keep 2 cards per page
    if (isOddCard) {
      const paddingNeeded = 33 - rowsUsedCount;
      if (paddingNeeded > 0) {
        currentRow += paddingNeeded;
      } else if (paddingNeeded < 0) {
        Logger.log(`Warning: Card ${successfulCards + 1} (${climber.name}) is ${-paddingNeeded} rows too tall`);
      }
      currentRow++;
    } else {
      const paddingNeeded = 34 - rowsUsedCount;
      if (paddingNeeded > 0) {
        currentRow += paddingNeeded;
      } else if (paddingNeeded < 0) {
        Logger.log(`Warning: Card ${successfulCards + 1} (${climber.name}) is ${-paddingNeeded} rows too tall`);
      }
      if (successfulCards < climbers.length - skippedClimbers.length - 1) {
        currentRow += 2;
      }
    }
    
    successfulCards++; // Increment successful card count
    
    } catch (error) {
      // Catch any unexpected errors
      skippedClimbers.push({
        name: climber.name,
        category: climber.category,
        reason: `Error: ${error.message}`
      });
      Logger.log(`Error creating card for ${climber.name}: ${error.message}`);
    }
  }
  
  return skippedClimbers;
}

function createCheckInLists(spreadsheet, climbers) {
  // Group climbers by session
  const sessionGroups = {};
  
  climbers.forEach(climber => {
    const session = climber.session;
    if (!sessionGroups[session]) {
      sessionGroups[session] = [];
    }
    sessionGroups[session].push(climber);
  });
  
  // Sort sessions numerically
  const sessions = Object.keys(sessionGroups).sort((a, b) => parseInt(a) - parseInt(b));
  
  // Create a check-in sheet for each session
  sessions.forEach(session => {
    const sheetName = `Session ${session} Check-in List`;
    
    // Delete existing sheet if it exists
    let checkInSheet = spreadsheet.getSheetByName(sheetName);
    if (checkInSheet) {
      spreadsheet.deleteSheet(checkInSheet);
    }
    
    checkInSheet = spreadsheet.insertSheet(sheetName);
    
    // Sort climbers in this session by category, then last name
    const sessionClimbers = sessionGroups[session].sort((a, b) => {
      const catCompare = compareCategoriesRedpoint(a.category, b.category);
      if (catCompare !== 0) return catCompare;
      return a.lastname.localeCompare(b.lastname);
    });
    
    // Create header row
    checkInSheet.getRange(1, 1).setValue('Name').setFontWeight('bold');
    checkInSheet.getRange(1, 2).setValue('Category').setFontWeight('bold');
    checkInSheet.getRange(1, 3).setValue('Session').setFontWeight('bold');
    checkInSheet.getRange(1, 4).setValue('Bib #').setFontWeight('bold');
    checkInSheet.getRange(1, 5).setValue('Checked In').setFontWeight('bold');
    
    // Set header row background
    checkInSheet.getRange(1, 1, 1, 5).setBackground('#d9d9d9');
    
    // Add climber data
    sessionClimbers.forEach((climber, index) => {
      const row = index + 2; // Start at row 2 (after header)
      
      checkInSheet.getRange(row, 1).setValue(climber.name);
      checkInSheet.getRange(row, 2).setValue(climber.category);
      checkInSheet.getRange(row, 3).setValue(climber.session);
      checkInSheet.getRange(row, 4).setValue(climber.bib);
      // Column 5 left blank for manual check-in marks
    });
    
    // Format columns
    checkInSheet.setColumnWidth(1, 200); // Name
    checkInSheet.setColumnWidth(2, 100); // Category
    checkInSheet.setColumnWidth(3, 100); // Session
    checkInSheet.setColumnWidth(4, 80);  // Bib #
    checkInSheet.setColumnWidth(5, 120); // Checked In
    
    // Add borders to all data
    const dataRange = checkInSheet.getRange(1, 1, sessionClimbers.length + 1, 5);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Freeze header row
    checkInSheet.setFrozenRows(1);
  });
}
