/* global document, Office */
var $ = require("jquery");
var CryptoJS = require("crypto-js");
var localStorage = require("localStorage");
var ApiURl = "https://dev.1viewtask.com/authapi/login/userLogin";
var rootURLCompliance = "https://dev.1viewtask.com/complianceapi/";
$("#lblUSerError").html(JSON.stringify(localStorage));
//document.getElementById("sideload-msg").style.display = "none";
document.getElementById("app-body").style.display = "flex";
document.getElementById("run").onclick = run;
document.getElementById("AddTask").onclick = CreateTask;

$("#Login-btn").bind("click", function (event) {
  event.preventDefault();
  {
    try {
      if ($("#userName").val() === "") {
        $("#item-Error").html(JSON.stringify("Please enter your email address"));
      }
      if ($("#passWord").val() === "") {
        $("#item-Error").html(JSON.stringify("Please enter your password"));
      }
      if ($("#userName").val() === "" && $("#passWord").val() === "") {
        $("#item-Error").html(JSON.stringify("Please provide the login details"));
      } else {
        getUserLogin();
        GetUserSettings();
      }
    } catch (error) {
      $("#item-Error").html(JSON.stringify("Inside Try Catch"));
    }
  }
});

GetUserSettings();

// document.ready(() => {
//     if (.host === Office.HostType.Outlook) {
//       // for checking response in Login Page
//     }
// });

export async function run() {
  /**
   * Insert your Outlook code here
   */
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  window.location.href = "new_task.html";

  Office.context.mailbox.userProfile.emailAddress;
  $("#userName").val(Office.context.mailbox.userProfile.emailAddress);
}

export async function CreateTask() {
  /**
   * Insert your Outlook code here
   */
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>start :</b> <br/>" + item.subject;
  try {
    var form = new FormData();
    form.append("companyID", "11");
    form.append("AccessType", "1");
    form.append("Client", "164");
    form.append("Location", "");
    form.append("Activity", "0");
    form.append("SubActivity", "");
    form.append("ComplianceType", "93");
    form.append("Section", "88");
    form.append("Priority", "3");
    form.append("ComplianceID", "0");
    form.append("ComplianceName", "1130-attached");
    form.append("Preparer", "52");
    form.append("Reviewer", "");
    form.append("Activitys", "Others-Others");
    form.append("ClientLocation", "Others-Registered");
    form.append("Esclation", "");
    form.append("PreparerAlert", "2");
    form.append("ReviewerAlert", "1");
    form.append("EsclationAlert", "2");
    form.append("PreparerAlertDurationType", "2");
    form.append("ReviewerAlertDurationType", "2");
    form.append("EscalationAlertDurationType", "2");
    form.append("preparerTaskDurationType", "2");
    form.append("reviewerTaskDurationType", "1");
    form.append("preparerDuration", "1");
    form.append("reviewerDuration", "15");
    form.append("editor", "This is Discription");
    form.append("Attachments", "[ ]");
    form.append(
      "SchedulersData",
      '{"frequencyID":"1","frequencyCode":"O","frequencyData":"{\\"OonTime\\":\\"2022-09-26 06:57:00\\"}"}--{"StartDate":"2022-09-26 06:57:00","EndDate":"2022-12-26 06:57:00","NoOfOcurrences":null}'
    );
    form.append("Penalty", "0");
    form.append("UserId", "52");
    form.append("CreatedBy", "");
    form.append("CreatedOn", "");
    form.append("complianceTags", "");

    // var settings = {
    //   "url": "https://qa.1viewtask.com/ComplianceMaster/CreateUpdateComplience",
    //   "type": "POST",
    //   "timeout": 2000,
    //   "processData": false,
    //   "mimeType": "multipart/form-data",
    //   "contentType": false,
    //   "data": form
    // };

    // $.ajax(settings).done(function (response) {
    //   console.log(response);
    //   document.getElementById("item-subject").innerHTML = "<b>success:</b> <br/>" + item.subject;
    // })
    // .fail(function (jqXHR, exception) {
    //   // Our error logic here
    //   document.getElementById("item-subject").innerHTML = "<b>Error:</b> <br/>" + JSON.stringify(jqXHR);

    // });

    // $.ajax({
    //     url: "http://localhost:49157/ComplianceMaster/CreateUpdateComplience",
    //     data: form,
    //     processData: false,
    //     contentType:false,
    //     type: 'POST',
    //     mimeType: 'multipart/form-data',
    //     dataType:'json',
    //     crossDomain: true,
    //     success: function (data) {
    //         console.log("success")
    //         document.getElementById("item-subject").innerHTML = "<b>success:</b> <br/>" + JSON.stringify(data) ;
    //     },
    //     error: function(error, textStatus,) {
    //         console.log("error" + error);
    //         document.getElementById("item-Error").innerHTML = "<b>eex error:</b> <br/>" + JSON.stringify(error);
    //     }
    //   });

    //////////////////// using fetch option
    var requestOptions = {
      method: "POST",
      body: form,
      redirect: "follow",
    };

    fetch("https://dev.1viewtask.com/ComplianceMaster/CreateUpdateComplience", requestOptions)
      .then((response) => response.text())
      .then((result) => {
        document.getElementById("item-subject").innerHTML = "<b>success:</b> <br/>" + result;
        //response.render('https://localhost:3000/new_task.html');
        window.location.href = "https://localhost:3000/new_task.html";
      })
      .catch((error) => {
        document.getElementById("item-Error").innerHTML = "<b>eex error:</b> <br/>" + JSON.parse(error);
      });
  } catch (error) {
    document.getElementById("item-Error").innerHTML = "<b>exception Tycatch:</b> <br/>" + error;
  }
}

export async function getUserLogin() {
  var lUserNAme = "";

  var userName = $("#userName").val();
  var passWord = $("#passWord").val();
  var dateStamp = new Date();
  var encodedTimeStamp = btoa(Date.parse(dateStamp) / 1000);
  var JsonDatetime = JSON.stringify(dateStamp);
  passWord = CryptoJS.AES.encrypt(passWord, encodedTimeStamp);

  var reqdata = {
    email: userName,
    password: passWord.toString(),
    loginDateTime: JsonDatetime.replace(/["']/g, ""),
  };
  // working
  // $("#lblUSerError").html(JSON.stringify(reqdata.password));
  $.ajax({
    url: "https://dev.1viewtask.com/authapi/login/userLogin",
    type: "POST",
    data: reqdata,
    dataType: "json",
    success: function (response, textStatus, xhr) {
      if (response.type == "failure") {
        if (textStatus == "error") {
          $("#lblUSerError").html(JSON.stringify("ErrorType : failure"));
        }
        $("#lblUSerError").html(JSON.stringify("No Record Found"));
      } else {
        // $("#lblUSerError").html(JSON.stringify(response.data.userId));

        // $('#lblUSerError').hide()
        // $('#sidebar-wrapper-container').show();

        // $("#lblUSerError").html(JSON.stringify(response.data.userId));

        lUserNAme += response.data.firstName + " " + response.data.lastName;
        localStorage.setItem("username", lUserNAme);
        localStorage.setItem("emailid", response.data.email);
        localStorage.setItem("UserID", response.data.userId);
        localStorage.setItem("IsSuperAdmin", response.data.isSuperAdmin);
        localStorage.setItem("isloggedin", true);
        localStorage.setItem("Token", response.data.token);
        localStorage.setItem("profilePicture", response.data.profilePictureFileName);
        // localStorage.setItem('SelectedCompany',response.data.SelectedCompany)

        //  Getting Response

        GetUserSettings(response.data.userId);

        // $("#lblUSerError").html(JSON.stringify('WORKINGS'));

        var x = 1;
        x = CheckSubscription(response.data.userId);
        if (x == 0) {
          $("#lblUSerError").html(JSON.stringify("WORKINGS"));
        } else {
          run();
          // $("#lblUSerError").html(JSON.stringify(localStorage.profilePicture));

          // $("#lblUSerError").html(JSON.stringify("Last Step"));
          // var baseUrl = "/Home/Index";
          // var baseUrl = "/MyActivity/Index";
          // window.location.href = baseUrl;
        }
      }
    },
    // error: function (xhr, textStatus, errorThrown) {
    //     console.log('Error in Operation');
    // }
  });
}

export async function GetUserSettings() {
  $.ajax({
    type: "Post",
    url: "https://dev.1viewtask.com/Home/GetUserSettings",
    headers: {
      Authorization: `Bearer ${localStorage.Token}`,
    },
    async: false,
    data: {
      userId: localStorage.UserID,
    },
    dataType: "json",
    success: function (response, textStatus, xhr) {
      var SelectedView = $.grep(Object(response), function (j) {
        return j.fieldName === "selectedview";
      });
      if (SelectedView == null || SelectedView == "") {
        localStorage.setItem("SelectedView", "kanbanview");
      } else {
        localStorage.setItem("SelectedView", SelectedView[0].fieldValue);
      }

      //setimezone from settings
      SelectedView = $.grep(Object(response), function (j) {
        return j.fieldName === "selectedtimezone";
      });
      if (SelectedView == null || SelectedView == "") {
        moment.tz.setDefault(moment.tz.guess()); // default will be machine timezone
        localStorage.timezone = moment().tz();
      } else {
        localStorage.timezone = SelectedView[0].fieldValue;
      }
    },
    error: function (xhr, textStatus, errorThrown) {
      console.log("Some issue with settings");
    },
  });
}

//###############################################
// export async function lsRememberMe(){
//   if (rmCheck.checked && emailInput.value !== "") {
//     localStorage.emailid = emailInput.value;
//     localStorage.checkbox = rmCheck.value;
//     localStorage.password = passinput.value;
//     } else {
//     localStorage.emailid = "";
//     localStorage.checkbox = "";
//     localStorage.password=""
//     }
// }

export async function CheckSubscription() {
  //alert(authtoken);
  var companylist = [];
  //debugger;
  $.ajax({
    url: "https://dev.1viewtask.com/Home/user/info",
    type: "POST",
    headers: {
      Authorization:
        "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOjMzOSwiaWF0IjoxNjY0NDM2MjAyfQ.5is7t5yICdxKT8_1WAp8l2MmT40cxc1wpnTzThpXeeA",
    },
    dataType: "json",
    data: {
      userId: localStorage.UserID,
      applicationId: 1,
    },
    success: function (response, textStatus, xhr) {
      //alert("success");
      $.each(response.data.companyUsers, function (data, val) {
        //debugger;
        companylist.push(val.masterCompany.companyId);
      });
    },
    error: function (xhr, textStatus, errorThrown) {
      //console.log(xhr);
      //alert(textStatus);
      console.log("Error in Operation");
    },
  });
  var FlagSubscription = 0;
  for (var i = 0; i < companylist.length; i++) {
    if (FlagSubscription == 1) {
      return;
    } else {
      var data = {
        storeProcedureName: "GetSubscriptionDetails",
        parameters: {
          companyid: companylist[i],
        },
      };
      //debugger;
      $.ajax({
        type: "POST",
        url: rootURLCompliance + "store/procedure/execute",
        data: JSON.stringify(data),
        dataType: "json",
        async: false,
        contentType: "application/json",
        success: function (resp) {
          $.each(resp.data.data[0], function (data, val) {
            //debugger;
            if (val.Subscription_ID > 0) {
              var today = new Date();
              if (val.Subscription_StartDate != null && val.Subscription_EndDate != null) {
                var Subscription_StartDateTime = new Date(val.Subscription_StartDate);
                var Subscription_EndDateTime = new Date(val.Subscription_EndDate);
                if (today > Subscription_StartDateTime && today < Subscription_EndDateTime) {
                  //alert("Pro");
                  FlagSubscription = 1;
                }
              }
              if (val.Trial_StartDate != null && val.Trial_EndDate != null) {
                var Trial_StartDateTime = new Date(val.Trial_StartDate);
                var Trial_EndDateTime = new Date(val.Trial_EndDate);
                if (today > Trial_StartDateTime && today < Trial_EndDateTime) {
                  //alert("trial on");
                  FlagSubscription = 1;
                }
              }
            }
          });
        },
      });
    }
  }
  return FlagSubscription;
}
// $("#lblUSerError").html(JSON.stringify("dgsds"));
