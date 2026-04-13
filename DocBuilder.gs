/**
 * Creates a Google Doc with a comments template for each course.
 * Each course gets its own section (separated by page breaks) with a table
 * listing students alphabetically with a blank comment column.
 *
 * @param {Array} coursesData - [{name: "Course Name", students: [{name, id}]}]
 * @param {string} gradingPeriod - e.g., "Fall Midterm", "Winter Term"
 * @returns {string} URL of the created Google Doc
 */
function createCommentsDoc(coursesData, gradingPeriod) {
  var docTitle = 'Comments — ' + gradingPeriod;
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();

  body.setFontFamily('Arial');

  for (var i = 0; i < coursesData.length; i++) {
    var course = coursesData[i];

    // Page break between courses (not before the first)
    if (i > 0) {
      body.appendPageBreak();
    }

    // Course heading
    var heading = body.appendParagraph(course.name + ' — ' + gradingPeriod);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    heading.setFontSize(16);
    heading.setBold(true);

    // Spacer
    body.appendParagraph('');

    if (course.students.length === 0) {
      body.appendParagraph('No students enrolled.').setItalic(true);
      continue;
    }

    // Build table data: header + student rows
    var tableData = [['Student Name', 'Comment']];
    for (var j = 0; j < course.students.length; j++) {
      tableData.push([course.students[j].name, '']);
    }

    var table = body.appendTable(tableData);

    // Style the header row
    var headerRow = table.getRow(0);
    for (var k = 0; k < headerRow.getNumCells(); k++) {
      var headerCell = headerRow.getCell(k);
      headerCell.setBackgroundColor('#4a86c8');
      headerCell.editAsText()
        .setForegroundColor('#ffffff')
        .setBold(true)
        .setFontSize(11);
      headerCell.setPaddingTop(6);
      headerCell.setPaddingBottom(6);
      headerCell.setPaddingLeft(8);
      headerCell.setPaddingRight(8);
    }

    // Style data rows with alternating colors
    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      var bgColor = (r % 2 === 0) ? '#f2f7fc' : '#ffffff';

      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        cell.setBackgroundColor(bgColor);
        cell.editAsText().setFontSize(10);
        cell.setPaddingTop(4);
        cell.setPaddingBottom(4);
        cell.setPaddingLeft(8);
        cell.setPaddingRight(8);
      }

      // Bold the student name
      row.getCell(0).editAsText().setBold(true);
    }

    // Student count footer
    var countText = body.appendParagraph(course.students.length + ' students');
    countText.setFontSize(9);
    countText.setForegroundColor('#888888');
    countText.setItalic(true);
  }

  doc.saveAndClose();
  return doc.getUrl();
}
