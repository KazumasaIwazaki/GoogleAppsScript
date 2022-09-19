function GetTimeStamp() {
  var date = new Date();
  var TimeStamp = Utilities.formatDate( date, 'Asia/Tokyo', 'yyyyMMdd_hhmmss');
  
  return TimeStamp;
}

function GetDate() {
  var date = new Date();
  var TimeStamp = Utilities.formatDate( date, 'Asia/Tokyo', 'yyyyMMdd');
  
  return TimeStamp;
}