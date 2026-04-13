/**
 * Adds the "Canvas Tools" menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Canvas Tools')
    .addItem('Set Canvas API Token', 'showTokenDialog')
    .addSeparator()
    .addItem('Generate Comments Template', 'showCourseSelector')
    .addToUi();
}

/**
 * Opens the token setup dialog.
 */
function showTokenDialog() {
  var html = HtmlService.createHtmlOutputFromFile('TokenDialog')
    .setWidth(450)
    .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Canvas API Setup');
}

/**
 * Saves the Canvas API token and base URL to user properties.
 */
function saveToken(token, baseUrl) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('CANVAS_TOKEN', token);
  props.setProperty('CANVAS_BASE_URL', baseUrl.replace(/\/+$/, ''));
}

/**
 * Fetches courses from Canvas and opens the course selector dialog.
 */
function showCourseSelector() {
  var props = PropertiesService.getUserProperties();
  var token = props.getProperty('CANVAS_TOKEN');
  var baseUrl = props.getProperty('CANVAS_BASE_URL');

  if (!token || !baseUrl) {
    SpreadsheetApp.getUi().alert(
      'Please set your Canvas API token first.\n\nGo to: Canvas Tools → Set Canvas API Token'
    );
    return;
  }

  try {
    var courses = getCourses(baseUrl, token);
    if (courses.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No active courses found.\n\nMake sure you are enrolled as a teacher in at least one Canvas course.'
      );
      return;
    }

    // Store courses temporarily for the dialog to read
    props.setProperty('TEMP_COURSES', JSON.stringify(courses));

    var html = HtmlService.createHtmlOutputFromFile('CourseSelector')
      .setWidth(500)
      .setHeight(520);
    SpreadsheetApp.getUi().showModalDialog(html, 'Generate Comments Template');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error fetching courses from Canvas:\n\n' + e.message);
  }
}

/**
 * Returns the temporarily stored courses list (called from CourseSelector dialog).
 */
function getStoredCourses() {
  var json = PropertiesService.getUserProperties().getProperty('TEMP_COURSES');
  return json ? JSON.parse(json) : [];
}

/**
 * Main function: fetches rosters and creates the Google Doc template.
 * Called from the CourseSelector dialog.
 */
function generateTemplate(selectedCourseIds, gradingPeriod) {
  var props = PropertiesService.getUserProperties();
  var token = props.getProperty('CANVAS_TOKEN');
  var baseUrl = props.getProperty('CANVAS_BASE_URL');
  var courses = JSON.parse(props.getProperty('TEMP_COURSES'));

  // Filter to only selected courses
  var selectedCourses = courses.filter(function (c) {
    return selectedCourseIds.indexOf(String(c.id)) > -1;
  });

  // Fetch roster for each selected course
  var coursesData = [];
  for (var i = 0; i < selectedCourses.length; i++) {
    var course = selectedCourses[i];
    var students = getRoster(baseUrl, token, course.id);
    coursesData.push({
      name: course.name,
      students: students
    });
  }

  // Create the Google Doc
  var docUrl = createCommentsDoc(coursesData, gradingPeriod);

  // Clean up temp data
  props.deleteProperty('TEMP_COURSES');

  return docUrl;
}
