function dummyData()
{
  // Before inserting, rename Sheet's name to "Employees_dev"

  const dummyData = [
    ['Name',     'Age', 'Join Date',           'Department',   'Score'],
    ['Alice',     29,   new Date(2021, 2, 10), 'Engineering',   87.5],
    ['Bob',       34,   new Date(2020, 6, 5),  'Marketing',     73.0],
    ['Charlie',   26,   new Date(2022, 0, 15), 'Sales',         90.2],
    ['Diana',     31,   new Date(2021, 9, 23), 'Engineering',   85.1],
    ['Ethan',     45,   new Date(2019, 3, 8),  'HR',            66.3],
    ['Fiona',     38,   new Date(2020, 11, 3), 'Finance',       92.0],
    ['George',    24,   new Date(2023, 4, 30), 'Sales',         88.7],
    ['Hannah',    29,   new Date(2022, 7, 19), 'Marketing',     77.4],
    ['Ian',       41,   new Date(2018, 1, 12), 'Engineering',   80.0],
    ['Jenny',     36,   new Date(2020, 10, 1), 'HR',            69.9],
    ['Kevin',     27,   new Date(2022, 2, 25), 'Finance',       93.5],
    ['Laura',     33,   new Date(2019, 6, 7),  'Sales',         82.1],
    ['Mike',      39,   new Date(2017, 8, 13), 'Engineering',   75.6],
    ['Nina',      28,   new Date(2021, 5, 11), 'HR',            70.0],
    ['Oscar',     30,   new Date(2023, 1, 2),  'Finance',       86.4],
    ['Pam',       37,   new Date(2020, 3, 16), 'Marketing',     91.2],
    ['Quincy',    32,   new Date(2019, 11, 9), 'Engineering',   83.3],
    ['Rachel',    25,   new Date(2022, 9, 29), 'Sales',         78.8],
    ['Sam',       40,   new Date(2016, 4, 5),  'HR',            68.5],
    ['Tina',      35,   new Date(2021, 6, 17), 'Marketing',     89.9],
    ['Uma',       29,   new Date(2023, 3, 21), 'Finance',       94.7]
  ];

  const sheet = SS.getSheetByName(SHEETNAME_EMPLOYEES);
  const headerRowCount = new SmartSheet(sheet).getHeaderRowCount();
  sheet.getRange(headerRowCount, 1, dummyData.length, dummyData[0].length).setValues(dummyData);

}