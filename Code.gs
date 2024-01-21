const Select_action ={
  check: 'http://koha.ekutuphane.gov.tr/cgi-bin/koha/opac-user.pl',
  renew:'/cgi-bin/koha/opac-renew.pl',
  logout:'http://koha.ekutuphane.gov.tr/cgi-bin/koha/opac-main.pl?logout.x=1',
  get_ics:'http://koha.ekutuphane.gov.tr/cgi-bin/koha/opac-ics.pl'
}
const Locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale().split('_')[0];
//max users =20
const Max_users = 20;
//max books per user =100
const Max_books = 100;
const List_names = Locale==='tr'? [['Kullanıcı adı'], ['İsim'],['İade tarihi'], ['Gün içinde'], ['Kontrol edilen son tarih']]:[['UserID'], ['Name'],['Date due'], ['In days'], ['Last checked']];
const Settings_names = Locale==='tr'? [['isim',	'kullanici adi', 'sifre', 'email']]:[['name',	'userid', 'pass', 'email']];
let Kitap_listesi_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Locale==='tr'?'Kitap listesi':'Book list');
let Settings_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('settings');
let App_sessions_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AppSessions');
function authorize(){
 UserProperties.setProperty('auth','true');
  console.log('authorized, running on open');
  onOpen();
}
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Koha App');
  let auth = UserProperties.getProperty('auth');
  if(auth !='true'){
    menu.addItem(Locale==='tr'?'Izin ver':'Authorize', 'authorize').addToUi();
  }
  else{
    menu.addItem(Locale==='tr'?'Yeni kullanici ekle':'Add new user', 'show_dialog')
    .addItem(Locale==='tr'?'Kullanıcı kayıtlarını al':'Check all users', 'check_all_users')
    .addToUi();
  }
  let curr_ss= SpreadsheetApp.getActiveSpreadsheet();

   if(Kitap_listesi_sheet === null){
    curr_ss.getActiveSheet().setName(Locale==='tr'?'Kitap listesi':'Book list');
    Kitap_listesi_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Locale==='tr'?'Kitap listesi':'Book list');
    Kitap_listesi_sheet.getRange('A1:A5').setValues(List_names).setFontWeight('bold');
    Kitap_listesi_sheet.getRange(1, 1, 5,Max_users+1).setNumberFormat('@STRING@');
  }
  if(Settings_sheet === null){
    curr_ss.insertSheet().setName('settings').hideSheet().getRange('A1:D1').setValues(Settings_names).setFontWeight('bold');
    Settings_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('settings');
  }

   if(App_sessions_sheet === null){
    curr_ss.insertSheet().setName('AppSessions').hideSheet();
    App_sessions_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AppSessions');
  }
}
function select_user_cookie(userid){
  var curr_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AppSessions');
 
  var all_userids_cookies = curr_sheet.getRange('A2:B').getValues().filter(String);
  
  var cookie_index = all_userids_cookies.map(itm => itm[0].toString()).indexOf(userid);
  var app_cookie = cookie_index  <0? '': all_userids_cookies[cookie_index][1];
  return app_cookie;
  
}
function get_cookie(url, userid, pwd){
  //logout first
  UrlFetchApp.fetch(Select_action.logout);
  var pld = { "userid":userid,   "password":pwd, "koha_login_context":"opac" };
  var options =
      {
        "method" : "post",
        "payload" : pld,
        "followRedirects" : false
      };
  var a=UrlFetchApp.fetch(url, options);
  var cont = a.getContentText();
  var resp= a.getResponseCode();
  var hdrs = a.getAllHeaders();
  
  var session_cookies = hdrs['Set-Cookie'];
  //Logger.log(session_cookies);
  //Logger.log(JSON.stringify(session_cookies));
  save_cookie(userid, session_cookies[0]);
  return session_cookies[0];
}
function save_cookie(userid, cookie){
  var user_id_colvals = App_sessions_sheet.getRange('A2:A').getValues().filter(String).map(itm => itm[0].toString());
  var userid_row = user_id_colvals.indexOf(userid);
  userid_row = userid_row<0? user_id_colvals.length: userid_row; 
  App_sessions_sheet.getRange(userid_row+2, 1,1,2).setValues([[userid, cookie]]);
  SpreadsheetApp.flush();
}
function get_pwd(userid){
  var all_userids = Settings_sheet.getRange('B2:C').getValues();
  var pwd_index = all_userids.map(itm => itm[0].toString()).indexOf(userid);
  var pwd = all_userids[pwd_index][1].toString();
  return pwd;
}
function get_name(userid){
  var all_data = Settings_sheet.getRange('A2:C').getValues();
  var person_index = all_data.map(itm => itm[1].toString()).indexOf(userid);
  var name = all_data[person_index][0].toString();
  return name;
}
function recheck_if_error(){
  var err_range = App_sessions_sheet.getRange('M2');
  var err = err_range.getValue();
  err_range.clear();
  if(err != ''){
    console.log('checking after error: '+err);
    check_all_users();
  }
}
function check_all_users(){
  var all_userids =Settings_sheet.getRange('B2:B').getValues().filter(String).map(itm => itm[0].toString());
  Kitap_listesi_sheet.getRange(1,2,Max_books,Max_users).clearContent();
  all_userids.forEach(itm => get_page_send_email(itm));
   
}
function get_email(userid){
  var all_data = Settings_sheet.getRange('A2:D').getValues();
  var email_index = all_data.map(itm => itm[1].toString()).indexOf(userid);
  var email = all_data[email_index][3].toString();
  return email;
}
function write_email(userid, titles, date_due, diff_days){
  var name = get_name(userid);
  var subject = name +": "+ titles.length.toString()+( titles.length>1?' books due in ':' book due in ' )+ diff_days.toString()+ ' days, on '+date_due;
  if(Locale === 'tr')
    subject = name +": "+ titles.length.toString()+' kitap '+ diff_days.toString()+ ' gün içinde '+date_due + 'tarihinde iade edilmeli';
  var colors = ['red','orangered', 'orange', 'gold', 'greenyellow', 'green','mediumturquoise', 'deepskyblue', 'blue', 'purple', 'darkviolet', 'deeppink'];
  var num_colors = colors.length;
  var html_titles = titles.reduce((prev,curr, currindex)=> prev+'<p><font size=3 color="'+colors[currindex % num_colors]+'">'+(currindex+1).toString()+') '+curr +'</font></p>', '' );
  var title= Locale==='tr'? diff_days +' gün içinde '+date_due+' tarihinde geri verilmesi gereken kitaplarınız: ':'Books due on '+date_due +' in '+diff_days+' days: ';
  var options = {
    htmlBody: '<p><font size=5><b><i>'+name +'</i></b></font></p><p><font size=5><b><i>'+title+'</i></b></font></p>' + html_titles
  };
  MailApp.sendEmail(get_email(userid),  subject, "", options);
}
function find_col(userid, sheet){
  //up to  max_users
  var first_row= sheet.getRange(1,1,1,Max_users+1).getValues()[0].filter(String);
  var col_index = first_row.indexOf(userid);
  col_index = col_index<0? sheet.getLastColumn()+1: col_index+1;
  return col_index;
}
function write_spreadsheet(userid, titles,dates, date_due, diff_days, last_checked){
  var name =get_name(userid);
  let last_col = Kitap_listesi_sheet.getLastColumn();
  if(last_col<1){
    Kitap_listesi_sheet.getRange('A1:A5').setValues(List_names).setFontWeight('bold');
    Kitap_listesi_sheet.getRange(1, 1, 5,Max_users+1).setNumberFormat('@STRING@');
  }
  var col = find_col(userid, Kitap_listesi_sheet);

  var array_to_write = [[userid,''],[name,''],[date_due,''], [diff_days,''], [last_checked,'']].concat(titles.map((itm, indx)=>[(indx+1).toString()+') '+ itm, dates[indx].substring(0,10)]));
  
  Kitap_listesi_sheet.getRange(1, col, titles.length +5, 2).setValues(array_to_write);

}

function get_page_send_email(userid){
  var cookie = select_user_cookie(userid);
  var url = Select_action.get_ics;
  try{
    var page  = UrlFetchApp.fetch(url,{"headers" : {"Cookie" : cookie } });
  }
  catch (e){
    console.log(e);
    App_sessions_sheet.getRange('M2').setValue(JSON.stringify(e));
    return;
  }
  var response_headers= page.getAllHeaders();
  var wrong_content_type = response_headers['Content-Type'] == "text/html; charset=utf-8";
  if(wrong_content_type){
    cookie=get_cookie(url,userid, get_pwd(userid));
    try{
     page  = UrlFetchApp.fetch(url,{"headers" : {"Cookie" : cookie} });
   }
   catch(e){
     console.log(e);
     App_sessions_sheet.getRange('M2').setValue(JSON.stringify(e));
    return;
   }

  }
  var today = new Date();
  var formatted_today= Utilities.formatDate(today,"GMT+2", 'MMM d, yy');
   var page_text=page.getContentText();
   var split_dtstart = page_text.split('DTSTART:').slice(1);
   var split_summary = split_dtstart.map(itm =itm =>itm.split('SUMMARY:'));
   var dates_invalid = split_summary.map(itm=>itm[0].trim().split('T'));
   var dates = dates_invalid.map(itm => itm[0].substring(0,4)+'-'+itm[0].substring(4,6)+'-'+itm[0].substring(6,8)+' '+itm[1].substring(0,2)+':'+itm[1].substring(2,4));
   var titles = split_summary.map(itm =>itm[1].split(' iade')[0]);
   var titles_to_return;
  if(titles.length>0){
   
   var unique_dates = dates.filter((item,index,arr) => arr.indexOf(item)===index).sort();
   var soonest_date = unique_dates[0];
   var soonest_as_date = new Date(soonest_date);
   
   var formatted_date= Utilities.formatDate(soonest_as_date,"GMT+2", 'EEE, MMMM d, yyyy');
   
   var diffTime = soonest_as_date -today;
   var diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)); 
   if(diffDays<3){
     titles_to_return = titles.filter((itm, indx)=>{if(dates[indx]==soonest_date){return itm}});
     write_email(userid, titles_to_return, formatted_date, diffDays);
    }
    write_spreadsheet(userid, titles,dates, formatted_date, diffDays, formatted_today);
  }
  else 
    write_spreadsheet(userid, [Locale==='tr'? 'Ödünç alınmış kitabınız yok': 'You don\'t have any borrowed books'], [''], '',0, formatted_today)
}

function show_dialog(){
  
  var obj = {
    title: Locale === 'tr'?'Lütfen yeni kullanıcı bilgilerini girin.': 'Please enter the info for the new user.',
    data:[
      {title:Locale === 'tr'? 'İsim: ': 'Name: ', id: 'name'},
      {title:Locale === 'tr'? 'Kullanıcı adı (TC kimlik numarası): ': 'User ID (TR identification number): ', id: 'userid'},
      {title: Locale === 'tr'? 'Şifre: ': 'Password', id:'pass'},
      {title: Locale === 'tr'? 'İade bilgisi için e-mail adresi: ': 'Email to receive return date notifications: ',id: 'email' }
    ],
    submit_func : 'record_info'
  };
  create_input_dialog(obj);
}
function record_info(obj){
  //get users, update or add new
  let new_userid = obj.userid;
  let new_row_values = [obj.name, new_userid, obj.pass, obj.email];
  let last_row = Settings_sheet.getLastRow();
  let settings_rng;
  let existing_users;
  console.log(JSON.stringify(obj));
  console.log(last_row);
  if(last_row>2){
    settings_rng = Settings_sheet.getRange('A2:D'+last_row.toString());
    existing_users = settings_rng.getDisplayValues();
    console.log(existing_users);
  }
  if(existing_users){
    let userid_arr = existing_users.map(itm =>itm[1]);
    console.log(userid_arr);
    let match_new_userid= userid_arr.indexOf(new_userid);
    if(match_new_userid>=0){
      existing_users[match_new_userid]=new_row_values;
      settings_rng.setValues(existing_users);
    }
    else{
      let new_row = (last_row+1).toString();
      Settings_sheet.getRange('A'+new_row+':D'+new_row).setValues([new_row_values])
    }
  }
  else{
    let new_row = (last_row+1).toString();
    Settings_sheet.getRange('A'+new_row+':D'+new_row).setValues([new_row_values])
  }
  
}

function create_input_dialog(input_obj){
  var html_string = '<html><body><p>'+ input_obj['title']+'</p><table>';
  var on_click = 'var post ={';
  var data = input_obj.data;
  for(var i=0; i<data.length; i++){
    let curr_id = data[i].id;
    if(curr_id === 'pass')
      html_string +="<tr><td>"+data[i].title +"</td><td><input type=\"password\" id=\""+curr_id+"\"></td></tr>";
    else if(curr_id === 'email')
      html_string +="<tr><td>"+data[i].title +"</td><td><input type=\"email\" id=\""+curr_id+"\"></td></tr>";
    else
      html_string +="<tr><td>"+data[i].title +"</td><td><input type=\"text\" id=\""+curr_id+"\"></td></tr>";
    on_click +=data[i].id+":$('#"+data[i].id+"').val()";
    if(i<data.length-1)
      on_click += ',';
  }
  on_click +="}; google.script.run."+input_obj.submit_func+"(post); setTimeout(function(){google.script.host.close();}, 100);"
  html_string+='<tr><td><button onclick="setTimeout(function(){google.script.host.close();}, 100);">Cancel</button></td><td><button onclick="'+on_click+'">Submit</button></td></tr></table></body><script src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script></html>';
  SpreadsheetApp.getUi().showDialog(HtmlService.createHtmlOutput(html_string));
}