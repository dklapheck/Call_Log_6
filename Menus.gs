function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teacher Tools')
    .addItem('Refresh Student Data', 'refreshStudentDataAndCounselors')
    .addSeparator()
    .addItem('Save SCC Call Entry', 'saveSccToRoster')
    .addItem('Log Selected SCC/Attempt in PowerSchool', 'logSccInPowerSchool')
    .addItem('Reset SCC Call Entry', 'resetCallEntry')
    .addSeparator()
    .addItem('Log Current ECC in PowerSchool', 'logCurrentEccRowAndOpenPowerSchool')
    .addSeparator()
    .addItem('Update Engageli Attendance', 'updateEngageliAttendance')
    .addItem('Set Up / Repair Attendance Settings', 'setupEngageliAttendanceSettings')
    .addSeparator()
    .addItem('Save WIG Snapshot', 'saveWigSnapshot')
    .addItem('Enable Tuesday WIG Snapshots', 'enableTuesdayWigSnapshots')
    .addToUi();
}
