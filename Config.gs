const APP_CONFIG = {
  sheets: {
    dash: 'Dash',
    onboarding: 'Onboarding',
    scc: 'SCC',
    ecc: 'ECC',
    callEntry: 'Call Entry',
    wig: 'WIG',
    studentData: 'StudentData',
    overrides: 'LocalOverrides',
    contacts: 'Contacts',
    star: 'STAR Data',
    settings: 'Instructions and Settings',
    automationLog: 'Automation Log'
  },

  source: {
    urlCell: 'H2',
    lastRefreshCell: 'H3',
    upstream: 'Upstream',
    proRoster: 'ProRoster',
    courseGrades: 'CourseGrades',
    addDrop: 'AddDrop',
    engageli: 'Engageli Data'
  },

  dash: {
    wigRow: 35
  }
};

const SCC_CONFIG = {
  sheets: {
    callEntry: APP_CONFIG.sheets.callEntry,
    roster: APP_CONFIG.sheets.scc
  },

  callEntry: {
    studentSelector: 'C2',
    studentId: 'O2',
    note: 'D1',
    toDo: 'I1',
    demographicsAction: 'B15',
    formRange: 'A7:I50',

    resetRanges: [
      'H14',
      'H15',
      'H17:I22',
      'C20',
      'C22',
      'C32:E32',
      'C36',
      'C37',
      'C39',
      'C40',
      'H42',
      'C47:H47'
    ]
  },

  roster: {
    headerRow: 1,
    studentIdHeader: 'Student Number',
    lastNameHeader: 'LAST NAME',
    firstNameHeader: 'FIRST NAME',
    completionHeader: 'SCC Completion',
    notesHeader: 'SCC Notes',
    toDoHeader: 'SCC To Do',
    completedValue: 'Completed'
  }
};
