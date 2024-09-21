const SHEET_NAME_PREFIX_REGEX = /^\d+\./;
const DATA_START_ROW = 6;
const DATA_END_ROW = 10000;
const COLUMN_HEADERS = ['No', 'Task', 'PIC', 'Status', 'Duration', 'Start Date', 'End Date', 'Skip', 'Skip Reason', 'Parallel', 'Issue Link', 'Skip Holidays', 'Dependency'];

// Define the statuses to exclude
const statusesToExclude = ['Done', 'Ready to Merge', 'Ready for Deployement', 'Ready to Test', 'Ready to Implement']; // Add more statuses as needed

// Convert the array of statuses to a string that can be used in the formula
const statusesList = statusesToExclude.map(status => `"${status}"`).join(",");


const TASKS_DATE_STATUS_RULES = [
  {
    formula: `=AND(ISNA(MATCH($D6, {${statusesList}}, 0)), $G6<TODAY(), NOT(ISBLANK($G6)))`,
    color: '#FF0000' // Red color for overdue dates
  },
  {
    formula: `=AND(ISNA(MATCH($D6, {${statusesList}}, 0)), $G6>=TODAY(), $G6<=TODAY()+3, NOT(ISBLANK($G6)))`,
    color: '#FFFF00' // Yellow color for dates within next 3 days
  }
];

const RESOURCE_DATE_RULES = [
  {
    formula: `=AND($G6<TODAY(), NOT(ISBLANK($G6)))`,
    color: '#FF0000' // Red color for overdue dates
  },
  {
    formula: `=AND($G6>=TODAY(), $G6<=TODAY()+3, NOT(ISBLANK($G6)))`,
    color: '#FFFF00' // Yellow color for dates within next 3 days
  }
];

// Define the status rules as an array of objects
const STATUS_RULES = [
  { status: 'Done', color: '#d9ead3' },
  { status: 'Testing Notes', color: '#ea9999' },
  { status: 'Ready to Implement', color: '#ffe599' },
  { status: 'Ready to Test', color: '#f1c232' },
  { status: 'In Progress', color: '#f9cb9c' }
];