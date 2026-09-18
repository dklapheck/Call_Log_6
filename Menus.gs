function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teacher Tools')
    .addItem('Refresh Student Data', 'refreshStudentData')
    .addSeparator()
    .addItem('Save SCC Call Entry', 'saveSccToRoster')
    .addItem('Reset SCC Call Entry', 'resetCallEntry')
    .addSeparator()
    .addItem('Save ECC Entry', 'saveEccEntry')
    .addItem('Save ECC Entry & Open PowerSchool', 'saveEccEntryAndOpenPowerSchool')
    .addItem('Archive Selected ECC Note & Open PowerSchool', 'archiveAndOpenPowerSchoolECC')
    .addSeparator()
    .addItem('Update Engageli Attendance', 'updateEngageliAttendance')
    .addItem('Set Up / Repair Attendance Settings', 'setupEngageliAttendanceSettings')
    .addSeparator()
    .addItem('Save WIG Snapshot', 'saveWigSnapshot')
    .addItem('Enable Tuesday WIG Snapshots', 'enableTuesdayWigSnapshots')
    .addToUi();
}
