/*
 * Lesson Grader web integration. Use this ONE file instead of the old
 * WebAppIntegration.gs + WebAppWorkflow.gs; do not keep either old copy.
 * Keep the existing Code.gs and WebApp.html. Exactly one doGet must exist.
 * This does not make existing spreadsheet-bound grading APIs web-safe.
 */
var LessonGraderWeb = (function () {
  var ROOT_NAME = 'Graded Assessments';
  var STAGES = ['Stage 4', 'Stage 5', 'Stage 6'];
  var ID_RE = /^[A-Za-z0-9_-]{20,200}$/;

  var SUB_HEADERS = [
    'SubmissionRecordID',
    'StudentUserID',
    'Status'
  ];

  var CONFIG_HEADERS = [
    'ConfigID',
    'CourseID',
    'CourseName',
    'CourseSection',
    'CourseWorkID',
    'AssignmentTitle',
    'MaxPoints',
    'SavedAt',
    'SavedBy',
    'Active',
    'TaskName',
    'DueDate',
    'AutoImported'
  ];

  var NON_TASK_SHEETS = {
    Submissions: true,
    ClassroomConfig: true,
    Setup: true,
    SubmissionFiles: true,
    CriteriaConfig: true,
    CriterionAssessments: true,
    AIAssessments: true,
    AssessmentHistory: true,
    ErrorLog: true,
    ClassLists: true,
    RubricProfiles: true,
    ApprovedGrades: true,
    Rubric: true,
    Criteria: true,
    Summary: true,
    Sheet1: true
  };

  function fail_(message) {
    return {
      success: false,
      message: message
    };
  }

  function validId_(id) {
    return typeof id === 'string' && ID_RE.test(id);
  }

  function ownRoots_() {
    return DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  }

  function header_(sheet, names) {
    if (
      !sheet ||
      sheet.getLastRow() < 1 ||
      sheet.getLastColumn() < names.length
    ) {
      return false;
    }

    var actual = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function (value) {
        return String(value).trim();
      });

    return names.every(function (name) {
      return actual.indexOf(name) !== -1;
    });
  }

  function validWorkbook_(ss) {
    return (
      header_(ss.getSheetByName(Config.SHEET_SUBMISSIONS), SUB_HEADERS) &&
      header_(
        ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG),
        CONFIG_HEADERS
      )
    );
  }

  function walk_(callback) {
    var roots = ownRoots_();

    while (roots.hasNext()) {
      var root = roots.next();

      for (var i = 0; i < STAGES.length; i++) {
        var stages = root.getFoldersByName(STAGES[i]);

        while (stages.hasNext()) {
          var courses = stages.next().getFolders();

          while (courses.hasNext()) {
            var categories = courses.next().getFolders();

            while (categories.hasNext()) {
              var files = categories
                .next()
                .getFilesByType(MimeType.GOOGLE_SHEETS);

              while (files.hasNext()) {
                callback(files.next());
              }
            }
          }
        }
      }
    }
  }

  function member_(id) {
    var found = false;

    walk_(function (file) {
      if (file.getId() === id) {
        found = true;
      }
    });

    return found;
  }

  function open_(id) {
    if (!validId_(id)) {
      throw new Error('Invalid workbook ID.');
    }

    var file = DriveApp.getFileById(id);

    if (
      file.isTrashed() ||
      file.getMimeType() !== MimeType.GOOGLE_SHEETS ||
      !member_(id)
    ) {
      throw new Error('Workbook is not in the Graded Assessments library.');
    }

    var ss = SpreadsheetApp.openById(id);

    if (!validWorkbook_(ss)) {
      throw new Error('Workbook is not initialised for Lesson Grader.');
    }

    return ss;
  }

  function taskNames_(ss) {
    return ss
      .getSheets()
      .filter(function (sheet) {
        return !NON_TASK_SHEETS[sheet.getName()] && !sheet.isSheetHidden();
      })
      .map(function (sheet) {
        return sheet.getName();
      })
      .sort();
  }

  function internalCategory_() {
    var explicit = PropertiesService
      .getScriptProperties()
      .getProperty('LESSON_GRADER_WEB_CATEGORY');

    if (explicit && String(explicit).trim()) {
      return String(explicit).trim();
    }

    var categories = {};
    var roots = ownRoots_();

    while (roots.hasNext()) {
      var root = roots.next();

      for (var i = 0; i < STAGES.length; i++) {
        var stages = root.getFoldersByName(STAGES[i]);

        while (stages.hasNext()) {
          var courses = stages.next().getFolders();

          while (courses.hasNext()) {
            var folders = courses.next().getFolders();

            while (folders.hasNext()) {
              categories[folders.next().getName()] = true;
            }
          }
        }
      }
    }

    var names = Object.keys(categories);

    return names.length === 1 ? names[0] : null;
  }

  function existingWorkbookId_(stage, courseCode, category, name) {
    var roots = ownRoots_();
    while (roots.hasNext()) {
      var stages = roots.next().getFoldersByName(stage);
      while (stages.hasNext()) {
        var courses = stages.next().getFoldersByName(courseCode);
        while (courses.hasNext()) {
          var categories = courses.next().getFoldersByName(category);
          while (categories.hasNext()) {
            var files = categories.next().getFilesByName(name);
            while (files.hasNext()) {
              var file = files.next();
              if (!file.isTrashed() && file.getMimeType() === MimeType.GOOGLE_SHEETS) return file.getId();
            }
          }
        }
      }
    }
    return '';
  }

  function list_() {
    try {
      var seen = {};
      var result = [];

      walk_(function (file) {
        var id = file.getId();

        if (seen[id] || file.isTrashed()) {
          return;
        }

        seen[id] = true;

        try {
          if (validWorkbook_(SpreadsheetApp.openById(id))) {
            result.push({
              id: id,
              name: file.getName()
            });
          }
        } catch (err) {
          console.error('Skipped inaccessible workbook: ' + err);
        }
      });

      result.sort(function (a, b) {
        var x = a.name.toLowerCase();
        var y = b.name.toLowerCase();

        if (x < y) return -1;
        if (x > y) return 1;
        if (a.id < b.id) return -1;
        if (a.id > b.id) return 1;
        return 0;
      });

      return {
        success: true,
        data: result
      };
    } catch (err) {
      console.error('Web workbook list: ' + err);
      return fail_('Unable to list grading workbooks.');
    }
  }

  function metadata_(id) {
    try {
      var ss = open_(id);
      var sheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var last = sheet.getLastRow();
      var count = 0;

      if (last > 1) {
        var values = sheet.getRange(2, 1, last - 1, 1).getValues();

        for (var i = 0; i < values.length; i++) {
          if (String(values[i][0] || '').trim()) {
            count++;
          }
        }
      }

      return {
        success: true,
        data: {
          id: id,
          name: ss.getName(),
          submissionCount: count
        }
      };
    } catch (err) {
      console.error('Web workbook metadata: ' + err);
      return fail_('Workbook is invalid, uninitialised, or inaccessible.');
    }
  }

  function rubrics_() {
    try {
      var data = [];
      var seen = {};
      var roots = ownRoots_();

      while (roots.hasNext()) {
        var shared = roots.next().getFoldersByName('Shared Rubrics');

        while (shared.hasNext()) {
          var files = shared.next().getFilesByType(MimeType.GOOGLE_SHEETS);

          while (files.hasNext()) {
            var file = files.next();

            if (!seen[file.getId()] && !file.isTrashed()) {
              seen[file.getId()] = true;

              data.push({
                id: file.getId(),
                name: file.getName()
              });
            }
          }
        }
      }

      data.sort(function (a, b) {
        return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
      });

      return {
        success: true,
        data: data
      };
    } catch (err) {
      console.error('Web rubric list: ' + err);
      return fail_('Unable to list shared rubrics.');
    }
  }

  function tasks_(id) {
    try {
      return {
        success: true,
        data: taskNames_(open_(id))
      };
    } catch (err) {
      console.error('Web task list: ' + err);
      return fail_('Unable to list tasks for this workbook.');
    }
  }

  function ensureRequiredSheet_(ss, sheetName, requiredHeaders) {
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet
        .getRange(1, 1, 1, requiredHeaders.length)
        .setValues([requiredHeaders]);
      sheet.setFrozenRows(1);
      return sheet;
    }

    var lastColumn = sheet.getLastColumn();
    var actual = lastColumn
      ? sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function (value) {
            return String(value).trim();
          })
      : [];

    var missing = requiredHeaders.filter(function (name) {
      return actual.indexOf(name) === -1;
    });

    if (missing.length) {
      var firstNewColumn = actual.length + 1;
      var requiredLastColumn = firstNewColumn + missing.length - 1;

      if (sheet.getMaxColumns() < requiredLastColumn) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          requiredLastColumn - sheet.getMaxColumns()
        );
      }

      sheet
        .getRange(1, firstNewColumn, 1, missing.length)
        .setValues([missing]);
    }

    return sheet;
  }

  function initialiseWorkbook_(workbookId) {
    var ss = SpreadsheetApp.openById(workbookId);

    ensureRequiredSheet_(
      ss,
      Config.SHEET_SUBMISSIONS,
      SUB_HEADERS
    );

    ensureRequiredSheet_(
      ss,
      Config.SHEET_CLASSROOM_CONFIG,
      CONFIG_HEADERS
    );

    SpreadsheetApp.flush();

    if (!validWorkbook_(ss)) {
      throw new Error(
        'Workbook initialization failed: required sheets or headers are missing.'
      );
    }

    return ss;
  }

  function create_(courseCode, taskName, year, rubricId, rubricName, document) {
    var code = String(courseCode || '').trim();
    var task = String(taskName || '').trim();

    var yr = String(
      year == null || String(year).trim() === ''
        ? new Date().getFullYear()
        : year
    ).trim();

    var category = internalCategory_();

    if (!category) {
      return fail_(
        'Internal category is not configured. Ask an administrator to set the LESSON_GRADER_WEB_CATEGORY script property to the existing category name.'
      );
    }

    if (
      !deriveStageFromCourseCode(code) ||
      !task ||
      /^\s*$/.test(task) ||
      !/^20\d{2}$/.test(yr)
    ) {
      return fail_('Enter a valid course code, task name and four-digit year.');
    }

    if (rubricId && rubricId !== '__UPLOAD__') {
      var rubrics = rubrics_();

      if (
        !rubrics.success ||
        !rubrics.data.some(function (rubric) {
          return rubric.id === rubricId;
        })
      ) {
        return fail_('Select an existing shared rubric.');
      }
    }

    if (
      rubricId === '__UPLOAD__' &&
      (!String(rubricName || '').trim() || !document)
    ) {
      return fail_('Upload a rubric document and enter its name.');
    }

    try {
      /* The setup helper may reuse an existing workbook and create a task
         tab. Back it up BEFORE calling that helper, not afterwards. */
      var expectedName = code + ' [' + yr + '] - ' + task + ' - Grading';
      var reusableId = existingWorkbookId_(deriveStageFromCourseCode(code), code, category, expectedName);
      if (reusableId) {
        DriveApp.getFileById(reusableId).makeCopy('[Backup ' + new Date().toISOString().replace(/[:.]/g, '-') + '] ' + expectedName);
      }
      var result = apiSubmitGradingWorkbookSetup(
        code,
        category,
        task,
        yr,
        rubricId || '',
        rubricName || '',
        document || null
      );

      if (!result || result.success !== true) {
        return result || fail_('Workbook setup returned no result.');
      }

      var workbookId = result.workbook && result.workbook.id;

      if (!validId_(workbookId)) {
        console.error(
          'Workbook setup succeeded without a valid workbook ID: ' +
          JSON.stringify(result)
        );

        return fail_(
          'Workbook was created, but setup did not return a valid workbook ID.'
        );
      }

      /* New files have no assessment history; reused files were backed up
         before the setup helper ran. */
      initialiseWorkbook_(workbookId);

      var ss = open_(workbookId);
      if (taskNames_(ss).indexOf(task) === -1) throw new Error('Expected task tab is missing: ' + task);
      WebGrading.prepare(ss, false);
      result.message += ' Review and activate a rubric for this task before grading. No rubric was activated automatically.';
      result.selectedRubricId = result.rubricFileId || '';
      return result;
    } catch (err) {
      console.error(
        'Web create workbook ' +
        (typeof workbookId !== 'undefined' ? workbookId : '') +
        ': ' +
        err.stack
      );

      return fail_('Workbook setup failed. Check Apps Script executions.');
    }
  }

  function link_(id, taskName, courseId, workId) {
    if (
      !validId_(id) ||
      !/^\d{1,40}$/.test(String(courseId || '')) ||
      !/^\d{1,40}$/.test(String(workId || ''))
    ) {
      return fail_('Invalid workbook or Classroom selection.');
    }

    var lock = LockService.getScriptLock();

    try {
      var ss = open_(id);
      var task = String(taskName || '').trim();

      if (taskNames_(ss).indexOf(task) === -1) {
        return fail_('Choose a task tab in the selected workbook.');
      }

      var course = Classroom.Courses.get(String(courseId));
      var coursework = Classroom.Courses.CourseWork.get(
        String(courseId),
        String(workId)
      );

      if (!coursework) {
        return fail_('Assignment not found.');
      }

      if (
        coursework.courseId &&
        String(coursework.courseId) !== String(courseId)
      ) {
        return fail_('Assignment does not belong to the selected course.');
      }

      var due = null;

      if (coursework.dueDate) {
        if (
          !coursework.dueDate.year ||
          !coursework.dueDate.month ||
          !coursework.dueDate.day
        ) {
          return fail_('Assignment has an invalid due date.');
        }

        var time = coursework.dueTime || {};

        due = new Date(Date.UTC(
          coursework.dueDate.year,
          coursework.dueDate.month - 1,
          coursework.dueDate.day,
          time.hours === undefined ? 23 : time.hours,
          time.minutes === undefined ? 59 : time.minutes
        ));

        if (isNaN(due.getTime())) {
          return fail_('Invalid Classroom due date.');
        }
      }

      var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);

      if (!header_(sheet, CONFIG_HEADERS)) {
        return fail_('Workbook ClassroomConfig schema is incomplete.');
      }

      lock.waitLock(10000);

      var values = sheet.getDataRange().getValues();
      var headers = values[0].map(function (value) {
        return String(value).trim();
      });

      var columns = {};

      CONFIG_HEADERS.forEach(function (header) {
        columns[header] = headers.indexOf(header);
      });

      for (var r = 1; r < values.length; r++) {
        if (
          String(values[r][columns.TaskName]) === task &&
          String(values[r][columns.CourseID]) === String(courseId) &&
          String(values[r][columns.CourseWorkID]) === String(workId)
        ) {
          return fail_('Assignment is already linked to this task.');
        }
      }

      var row = new Array(headers.length).fill('');

      row[columns.ConfigID] = Utilities.getUuid();
      row[columns.CourseID] = String(courseId);
      row[columns.CourseName] = course.name || '';
      row[columns.CourseSection] = course.section || '';
      row[columns.CourseWorkID] = String(workId);
      row[columns.AssignmentTitle] = coursework.title || '';
      row[columns.MaxPoints] =
        coursework.maxPoints == null ? 0 : coursework.maxPoints;
      row[columns.SavedAt] = new Date().toISOString();
      row[columns.SavedBy] =
        Session.getEffectiveUser().getEmail() || '';
      row[columns.Active] = true;
      row[columns.TaskName] = task;
      row[columns.DueDate] = due ? due.toISOString() : '';
      row[columns.AutoImported] = false;

      sheet.appendRow(row);

      var message = due
        ? 'Linked “' + (coursework.title || 'Untitled assignment') +
          '” to “' + task +
          '”. Hourly auto-import trigger must be configured separately.'
        : 'Assignment linked; automatic due-date import is unavailable. Manual import has not been verified.';

      return {
        success: true,
        message: message
      };
    } catch (err) {
      console.error('Web Classroom link: ' + err);

      return fail_(
        'Unable to link assignment. Check Classroom access and workbook edit permission.'
      );
    } finally {
      try {
        lock.releaseLock();
      } catch (ignore) {}
    }
  }

  function getTaskAssessmentOverview_(spreadsheetId, taskName) {
    try {
      var ss = open_(spreadsheetId);
      var task = String(taskName || '').trim();

      if (!task || taskNames_(ss).indexOf(task) === -1) {
        return fail_('Choose a valid visible task tab.');
      }

      var warnings = [];
      var rubricRows = [];
      var submissions = [];

      var rubricSheet = ss.getSheetByName('Rubric');

      if (rubricSheet && rubricSheet.getLastRow() > 1) {
        var rubricValues = rubricSheet.getDataRange().getValues();
        var rubricHeaders = rubricValues[0].map(function (header) {
          return String(header || '').trim();
        });

        var criterionIndex = rubricHeaders.indexOf('Criterion');
        var partIndex = rubricHeaders.indexOf('Part');
        var sectionIndex = rubricHeaders.indexOf('Section');
        var maxMarksIndex = rubricHeaders.indexOf('MaxMarks');
        var outcomeIndex = rubricHeaders.indexOf('Outcome');
        var bandIndex = rubricHeaders.indexOf('Band');
        var descriptionIndex = rubricHeaders.indexOf('Description');

        if (
          criterionIndex === -1 ||
          bandIndex === -1 ||
          descriptionIndex === -1
        ) {
          warnings.push(
            'The Rubric sheet does not contain the expected Criterion, Band and Description columns.'
          );
        } else {
          for (var r = 1; r < rubricValues.length; r++) {
            var row = rubricValues[r];
            var criterion = String(row[criterionIndex] || '').trim();

            if (!criterion) {
              continue;
            }

            rubricRows.push({
              criterion: criterion,
              part: partIndex === -1 ? '' : row[partIndex],
              section: sectionIndex === -1 ? '' : row[sectionIndex],
              maxMarks: maxMarksIndex === -1 ? '' : row[maxMarksIndex],
              outcome: outcomeIndex === -1 ? '' : row[outcomeIndex],
              band: bandIndex === -1 ? '' : row[bandIndex],
              description: descriptionIndex === -1 ? '' : row[descriptionIndex]
            });
          }
        }
      } else {
        warnings.push(
          'No human-facing Rubric sheet was found, or it does not contain rubric rows.'
        );
      }

      var submissionsSheet = ss.getSheetByName(
        Config.SHEET_SUBMISSIONS
      );

      if (!submissionsSheet || submissionsSheet.getLastRow() < 1) {
        warnings.push('No Submissions sheet was found in this workbook.');
      } else {
        var submissionValues = submissionsSheet.getDataRange().getValues();
        var submissionHeaders = submissionValues[0].map(function (header) {
          return String(header || '').trim();
        });

        function submissionColumn_(name) {
          return submissionHeaders.indexOf(name);
        }

        var idIndex = submissionColumn_('SubmissionRecordID');
        var taskIndex = submissionColumn_('Task');
        var taskNameIndex = submissionColumn_('TaskName');
        var studentNameIndex = submissionColumn_('StudentName');
        var studentEmailIndex = submissionColumn_('StudentEmail');
        var statusIndex = submissionColumn_('Status');
        var classroomStateIndex = submissionColumn_('ClassroomState');
        var turnedInTimeIndex = submissionColumn_('TurnedInTime');
        var updateTimeIndex = submissionColumn_('UpdateTime');
        var lateIndex = submissionColumn_('Late');
        var currentOfficialIndex = submissionColumn_('CurrentOfficial');
        var approvedAtIndex = submissionColumn_('ApprovedAt');
        var lockedAtIndex = submissionColumn_('LockedAt');
        var classroomGradeIndex = submissionColumn_(
          'ClassroomAssignedGrade'
        );

        if (idIndex === -1) {
          warnings.push(
            'The Submissions sheet does not contain a SubmissionRecordID column.'
          );
        } else {
          for (var s = 1; s < submissionValues.length; s++) {
            var submissionRow = submissionValues[s];
            var rowTask = '';

            if (taskIndex !== -1) {
              rowTask = String(submissionRow[taskIndex] || '').trim();
            } else if (taskNameIndex !== -1) {
              rowTask = String(
                submissionRow[taskNameIndex] || ''
              ).trim();
            }

            if (rowTask !== task) {
              continue;
            }

            var submissionId = String(
              submissionRow[idIndex] || ''
            ).trim();

            if (!submissionId) {
              continue;
            }

            submissions.push({
              submissionRecordId: submissionId,
              studentName:
                studentNameIndex === -1
                  ? ''
                  : submissionRow[studentNameIndex],
              studentEmail:
                studentEmailIndex === -1
                  ? ''
                  : submissionRow[studentEmailIndex],
              status:
                statusIndex === -1
                  ? ''
                  : submissionRow[statusIndex],
              classroomState:
                classroomStateIndex === -1
                  ? ''
                  : submissionRow[classroomStateIndex],
              turnedInTime:
                turnedInTimeIndex === -1
                  ? ''
                  : submissionRow[turnedInTimeIndex],
              updateTime:
                updateTimeIndex === -1
                  ? ''
                  : submissionRow[updateTimeIndex],
              late:
                lateIndex === -1
                  ? ''
                  : submissionRow[lateIndex],
              currentOfficial:
                currentOfficialIndex === -1
                  ? ''
                  : submissionRow[currentOfficialIndex],
              approvedAt:
                approvedAtIndex === -1
                  ? ''
                  : submissionRow[approvedAtIndex],
              lockedAt:
                lockedAtIndex === -1
                  ? ''
                  : submissionRow[lockedAtIndex],
              currentMark:
                classroomGradeIndex === -1
                  ? ''
                  : submissionRow[classroomGradeIndex]
            });
          }
        }
      }

      submissions.sort(function (a, b) {
        var aName = String(a.studentName || '').toLowerCase();
        var bName = String(b.studentName || '').toLowerCase();

        if (aName < bName) {
          return -1;
        }

        if (aName > bName) {
          return 1;
        }

        return String(b.updateTime || '').localeCompare(
          String(a.updateTime || '')
        );
      });

      return {
        success: true,
        data: {
          workbook: {
            id: spreadsheetId,
            name: ss.getName()
          },
          taskName: task,
          rubric: {
            available: rubricRows.length > 0,
            rows: rubricRows
          },
          submissions: submissions,
          warnings: warnings
        }
      };
    } catch (err) {
      console.error(
        'Web task assessment overview: ' +
        (err && err.stack ? err.stack : err)
      );

      return fail_(
        'Unable to load the read-only task assessment overview.'
      );
    }
  }

  function getSubmissionAssessmentDetail_(
    spreadsheetId,
    taskName,
    submissionRecordId
  ) {
    try {
      var ss = open_(spreadsheetId);
      var task = String(taskName || '').trim();
      var submissionId = String(submissionRecordId || '').trim();

      if (!task || taskNames_(ss).indexOf(task) === -1) {
        return fail_('Choose a valid visible task tab.');
      }

      if (!submissionId) {
        return fail_('Choose a valid student submission.');
      }

      var warnings = [];
      var submissionDetail = null;
      var attachments = [];
      var criterionStates = [];
      var history = [];

      var submissionsSheet = ss.getSheetByName(
        Config.SHEET_SUBMISSIONS
      );

      if (!submissionsSheet || submissionsSheet.getLastRow() < 1) {
        return fail_('The Submissions sheet is unavailable.');
      }

      var submissionValues = submissionsSheet.getDataRange().getValues();
      var submissionHeaders = submissionValues[0].map(function (header) {
        return String(header || '').trim();
      });

      function submissionIndex_(name) {
        return submissionHeaders.indexOf(name);
      }

      var idIndex = submissionIndex_('SubmissionRecordID');
      var taskIndex = submissionIndex_('Task');
      var taskNameIndex = submissionIndex_('TaskName');

      if (idIndex === -1) {
        return fail_(
          'The Submissions sheet does not contain a SubmissionRecordID column.'
        );
      }

      for (var r = 1; r < submissionValues.length; r++) {
        var row = submissionValues[r];

        if (String(row[idIndex] || '').trim() !== submissionId) {
          continue;
        }

        var rowTask = '';

        if (taskIndex !== -1) {
          rowTask = String(row[taskIndex] || '').trim();
        } else if (taskNameIndex !== -1) {
          rowTask = String(row[taskNameIndex] || '').trim();
        }

        if (rowTask !== task) {
          continue;
        }

        function field_(name) {
          var index = submissionHeaders.indexOf(name);
          return index === -1 ? '' : row[index];
        }

        submissionDetail = {
          submissionRecordId: submissionId,
          studentName: field_('StudentName'),
          studentEmail: field_('StudentEmail'),
          taskName: rowTask,
          status: field_('Status'),
          classroomState: field_('ClassroomState'),
          turnedInTime: field_('TurnedInTime'),
          updateTime: field_('UpdateTime'),
          late: field_('Late'),
          currentOfficial: field_('CurrentOfficial'),
          approvedAt: field_('ApprovedAt'),
          lockedAt: field_('LockedAt'),
          currentMark: field_('ClassroomAssignedGrade')
        };

        break;
      }

      if (!submissionDetail) {
        return fail_(
          'The selected submission was not found for this task.'
        );
      }

      var filesSheet = ss.getSheetByName(
        Config.SHEET_SUBMISSION_FILES
      );

      if (filesSheet && filesSheet.getLastRow() > 1) {
        var fileValues = filesSheet.getDataRange().getValues();
        var fileHeaders = fileValues[0].map(function (header) {
          return String(header || '').trim();
        });

        function fileIndex_(name) {
          return fileHeaders.indexOf(name);
        }

        var fileSubmissionIdIndex = fileIndex_('SubmissionRecordID');

        if (fileSubmissionIdIndex === -1) {
          warnings.push(
            'SubmissionFiles does not contain a SubmissionRecordID column.'
          );
        } else {
          for (var f = 1; f < fileValues.length; f++) {
            var fileRow = fileValues[f];

            if (
              String(fileRow[fileSubmissionIdIndex] || '').trim() !==
              submissionId
            ) {
              continue;
            }

            function fileField_(name) {
              var index = fileHeaders.indexOf(name);
              return index === -1 ? '' : fileRow[index];
            }

            attachments.push({
              fileName: fileField_('FileName'),
              mimeType: fileField_('MimeType'),
              alternateLink: fileField_('AlternateLink'),
              thumbnailUrl: fileField_('ThumbnailUrl'),
              fileSize: fileField_('FileSize'),
              sourceType: fileField_('SourceType'),
              eligibleForAI: fileField_('EligibleForAI'),
              aiReviewStatus: fileField_('AIReviewStatus'),
              limitations: fileField_('Limitations')
            });
          }
        }
      } else {
        warnings.push(
          'No SubmissionFiles sheet was found, or it has no attachment rows.'
        );
      }

      var criterionSheet = ss.getSheetByName(
        'CriterionAssessments'
      );

      if (criterionSheet && criterionSheet.getLastRow() > 1) {
        var criterionValues = criterionSheet.getDataRange().getValues();
        var criterionHeaders = criterionValues[0].map(function (header) {
          return String(header || '').trim();
        });

        function criterionIndex_(name) {
          return criterionHeaders.indexOf(name);
        }

        var criterionSubmissionIdIndex = criterionIndex_(
          'SubmissionRecordID'
        );

        if (criterionSubmissionIdIndex === -1) {
          warnings.push(
            'CriterionAssessments does not contain a SubmissionRecordID column.'
          );
        } else {
          for (var c = 1; c < criterionValues.length; c++) {
            var criterionRow = criterionValues[c];

            if (
              String(
                criterionRow[criterionSubmissionIdIndex] || ''
              ).trim() !== submissionId
            ) {
              continue;
            }

            function criterionField_(name) {
              var index = criterionHeaders.indexOf(name);
              return index === -1 ? '' : criterionRow[index];
            }

            criterionStates.push({
              criterionId: criterionField_('CriterionID'),
              criterionTitle: criterionField_('CriterionTitle'),
              band:
                criterionField_('SelectedBand') ||
                criterionField_('Band'),
              state:
                criterionField_('State') ||
                criterionField_('AssessmentState'),
              assessmentState:
                criterionField_('AssessmentStatus') ||
                criterionField_('ApprovalState') ||
                criterionField_('Status'),
              teacherNotes:
                criterionField_('TeacherNotes') ||
                criterionField_('Notes') ||
                criterionField_('Feedback'),
              aiDraftState:
                criterionField_('AIDraftState') ||
                criterionField_('AIState') ||
                criterionField_('AISuggestion'),
              evidenceSummary:
                criterionField_('EvidenceSummary') ||
                criterionField_('Evidence'),
              evidenceLocation:
                criterionField_('EvidenceLocation') ||
                criterionField_('EvidenceReference'),
              source:
                criterionField_('Source') ||
                criterionField_('UpdatedBy')
            });
          }
        }
      } else {
        warnings.push(
          'No CriterionAssessments sheet was found, or it has no matching rows.'
        );
      }

      var historySheet = ss.getSheetByName(
        'AssessmentHistory'
      );

      if (historySheet && historySheet.getLastRow() > 1) {
        var historyValues = historySheet.getDataRange().getValues();
        var historyHeaders = historyValues[0].map(function (header) {
          return String(header || '').trim();
        });

        function historyIndex_(name) {
          return historyHeaders.indexOf(name);
        }

        var historySubmissionIdIndex = historyIndex_(
          'SubmissionRecordID'
        );

        if (historySubmissionIdIndex === -1) {
          warnings.push(
            'AssessmentHistory does not contain a SubmissionRecordID column.'
          );
        } else {
          for (var h = 1; h < historyValues.length; h++) {
            var historyRow = historyValues[h];

            if (
              String(
                historyRow[historySubmissionIdIndex] || ''
              ).trim() !== submissionId
            ) {
              continue;
            }

            function historyField_(name) {
              var index = historyHeaders.indexOf(name);
              return index === -1 ? '' : historyRow[index];
            }

            history.push({
              timestamp:
                historyField_('Timestamp') ||
                historyField_('CreatedAt') ||
                historyField_('UpdatedAt'),
              action:
                historyField_('Action') ||
                historyField_('Event'),
              state:
                historyField_('State') ||
                historyField_('AssessmentState'),
              author:
                historyField_('Author') ||
                historyField_('ChangedBy') ||
                historyField_('CreatedBy'),
              assessmentId:
                historyField_('AssessmentID') ||
                historyField_('AssessmentRecordID'),
              version:
                historyField_('Version') ||
                historyField_('AssessmentVersion')
            });
          }
        }
      } else {
        warnings.push(
          'No AssessmentHistory sheet was found, or it has no matching rows.'
        );
      }

      return {
        success: true,
        data: {
          submission: submissionDetail,
          attachments: attachments,
          criterionStates: criterionStates,
          history: history,
          warnings: warnings
        }
      };
    } catch (err) {
      console.error(
        'Web submission assessment detail: ' +
        (err && err.stack ? err.stack : err)
      );

      return fail_(
        'Unable to load the read-only submission assessment detail.'
      );
    }
  }

  function isSharedRubric_(id) {
    if (!validId_(id)) return false;
    var listing = rubrics_();
    return listing.success && listing.data.some(function (item) { return item.id === id; });
  }

  return {
    list: list_,
    metadata: metadata_,
    rubrics: rubrics_,
    tasks: tasks_,
    create: create_,
    link: link_,
    open: open_,
    taskNames: taskNames_,
    isSharedRubric: isSharedRubric_,
    getTaskAssessmentOverview: getTaskAssessmentOverview_,
    getSubmissionAssessmentDetail: getSubmissionAssessmentDetail_
  };
})();

function apiWebListCandidateWorkbooks() {
  WebAccess.requireTeacher();
  return LessonGraderWeb.list();
}

function apiWebGetWorkbookMetadata(spreadsheetId) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.metadata(spreadsheetId);
}

function apiWebListSharedRubrics() {
  WebAccess.requireTeacher();
  return LessonGraderWeb.rubrics();
}

function apiWebListWorkbookTasks(spreadsheetId) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.tasks(spreadsheetId);
}

function apiWebCreateWorkbook(
  courseCode,
  taskName,
  year,
  rubricSelection,
  rubricName,
  rubricDocument
) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.create(
    courseCode,
    taskName,
    year,
    rubricSelection,
    rubricName,
    rubricDocument
  );
}

function apiWebLinkClassroomAssignment(
  spreadsheetId,
  taskName,
  courseId,
  courseWorkId
) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.link(
    spreadsheetId,
    taskName,
    courseId,
    courseWorkId
  );
}

function apiWebGetTaskAssessmentOverview(spreadsheetId, taskName) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.getTaskAssessmentOverview(
    spreadsheetId,
    taskName
  );
}

function apiWebGetSubmissionAssessmentDetail(
  spreadsheetId,
  taskName,
  submissionRecordId
) {
  WebAccess.requireTeacher();
  return LessonGraderWeb.getSubmissionAssessmentDetail(
    spreadsheetId,
    taskName,
    submissionRecordId
  );
}

/* Read-only HTTP handshake for the GitHub Pages launcher (site/index.html).
 * The launcher cannot use google.script.run, so it needs one tiny cross-origin
 * answer: "is this deployment alive and is the signed-in teacher allowlisted?".
 * Deliberately narrow:
 *   - an allowlist of probe names, never a caller-chosen function name;
 *   - WebAccess.probe() returns no key, no allowlist and only a masked email;
 *   - the JSONP callback name is restricted to a plain JS identifier;
 *   - any failure still returns a normal, safe response instead of a stack trace.
 * Grading, import and settings calls stay inside the Apps Script-served app. */
var LessonGraderLauncherApi = (function () {
  var PARAMETER = 'lgapi';
  var SCHEMA = '1';
  var PROBES = { status: true };
  var CALLBACK_RE = /^[A-Za-z_$][A-Za-z0-9_$]{0,63}$/;

  function requested(e) {
    var params = (e && e.parameter) || {};
    return String(params[PARAMETER] || '').trim().toLowerCase();
  }

  function callbackName(params) {
    var name = String((params && params.callback) || '').trim();
    return CALLBACK_RE.test(name) ? name : '';
  }

  function bodyFor(call) {
    if (!Object.prototype.hasOwnProperty.call(PROBES, call)) {
      return { success: false, message: 'Unknown launcher probe "' + call + '". This endpoint is read-only.' };
    }
    return WebAccess.probe();
  }

  function respond(call, params) {
    var body;
    try {
      body = bodyFor(call);
    } catch (err) {
      body = { success: false, message: 'The launcher probe could not be answered.' };
    }
    body.probe = call;
    body.schema = SCHEMA;
    var callback = callbackName(params);
    var json = JSON.stringify(body);
    if (!callback) {
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return { PARAMETER: PARAMETER, requested: requested, respond: respond };
})();

/* HTML partials the web app shell may include. Closed list: include() is a
 * global, so it must never return arbitrary project files. */
var WEB_PARTIALS = { WebAppStyles: true, WebAppClient: true };

function include(name) {
  if (!Object.prototype.hasOwnProperty.call(WEB_PARTIALS, name)) throw new Error('Unknown partial.');
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/*
 * Keep this as the only doGet in the entire Apps Script project.
 * A request without the launcher parameter serves the web app.
 * HtmlService ignores <meta name="viewport"> inside the file, so the viewport
 * is added here (REVIEW.md F05). Framing by other sites stays blocked unless
 * the owner sets LESSON_GRADER_ALLOW_EMBED=true for the GitHub launcher (F34).
 */
function doGet(e) {
  var call = LessonGraderLauncherApi.requested(e);
  if (call) return LessonGraderLauncherApi.respond(call, (e && e.parameter) || {});
  var output = HtmlService.createTemplateFromFile('WebApp').evaluate()
    .setTitle('Lesson Grader')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (AppConfig.flag(AppConfig.PROPS.ALLOW_EMBED)) {
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return output;
}
