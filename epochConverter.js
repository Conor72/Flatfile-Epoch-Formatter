//Converts several date formats to epoch
//Invalid dates will be flagged
//Dates in the past for 'CUSTOMHEADER1' will be flagged

const dfns = require("date-fns");
async function ExcelDateToJSDate(serial) {      //Excel converters
  var utc_days = Math.floor(serial - 25568);
  var utc_value = utc_days * 86400;
  var date_info = new Date(utc_value * 1000);
  var fractional_day = serial - Math.floor(serial) + 0.0000001;
  var total_seconds = Math.floor(86400 * fractional_day);
  var seconds = total_seconds % 60;
  total_seconds -= seconds;
  var hours = Math.floor(total_seconds / (60 * 60));
  var minutes = Math.floor(total_seconds / 60) % 60;
  return new Date(
    date_info.getFullYear(),
    date_info.getMonth(),
    date_info.getDate(),
    hours,
    minutes,
    seconds
  );
}
function dateIsValid(date) {
  return date instanceof Date && !isNaN(date);
}
module.exports = async ({ recordBatch, session, logger }) => {
  await Promise.all(
    await recordBatch.records.map(async (record) => {
      
      

////////// CUSTOMHEADER1 //////////

//Replace CUSTOMERHEADER1 with the name of header on spreadsheet

      if (record.get("CUSTOMHEADER1") !== "") {        
        if (isNaN(record.get("CUSTOMHEADER1"))) {
          var stringdate = new Date(record.get("CUSTOMHEADER1"));  //Grabs all cells in the CUSTOMHEADER1 column
          var epochDate = new Date(stringdate / 1000);          //Epoch converter to seconds NOT miliseconds
          record.set("CUSTOMHEADER1", "INVALID DATE");             //Pastes INVALID DATE error into column for incorrect dates

          if (dateIsValid(epochDate)) {
            var today = new Date();
            logger.info("Todays Date:" + today);
            logger.info("Expiry Date:" + stringdate);
            if (stringdate < today) {
              //If CUSTOMHEADER1 is before today...
              record.set("CUSTOMHEADER1", "Date In The Past");     //Instead of converting to epoch will leave error message
              logger.info("Date In The Past");
            } else if (stringdate > today) {                    //If CUSTOMHEADER1 is in the future...
              record.set("CUSTOMHEADER1", epochDate.valueOf());    //Goes ahead and converts date to epoch format
              logger.info("Correct - Date in the future");
            }
          } else {
            record.addError("CUSTOMHEADER1", "Not a valid date");
          }
        } else {
          var exceldate = await ExcelDateToJSDate(record.get("CUSTOMHEADER1"));  // Excel value coming in
          if (dateIsValid(exceldate)) {
            record.set("CUSTOMHEADER1", dfns.format(exceldate, "yyyy-MM-dd"));
          } else {
            record.addError("CUSTOMHEADER1", "Not a valid date");
          }
        }
      }

////////// CUSTOMHEADER2 //////////

//Replace CUSTOMERHEADER2 with the name of header on spreadsheet

      if (record.get("CUSTOMHEADER2") !== "") {                 //If CUSTOMHEADER2 cells are not Empty
        if (isNaN(record.get("CUSTOMHEADER2"))) {
          //String date, no change needed
          var stringdate = new Date(record.get("CUSTOMHEADER2"));
          var epochDate = new Date(stringdate / 1000);
          record.set("CUSTOMHEADER2", "INVALID DATE");          //Pastes INVALID DATE error into column for wrong dates
          if (dateIsValid(epochDate)) {
            record.set("CUSTOMHEADER2", epochDate.valueOf());
          } else {
            record.addError("CUSTOMHEADER2", "Not a valid date");
          }
        } else {
          var exceldate = await ExcelDateToJSDate(record.get("CUSTOMHEADER2")); // Excel value coming in
          if (dateIsValid(exceldate)) {
            record.set("CUSTOMHEADER2", dfns.format(exceldate, "yyyy-MM-dd"));
          } else {
            record.addError("CUSTOMHEADER2", "Not a valid date");
          }
        }
      }
    })
  );
};
