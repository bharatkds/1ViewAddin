/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var localStorage = require("localStorage");
var alert = require("jquery-confirm");
var $ = require("jquery");

// URL
var rootURLCompliance = "https://dev.1viewtask.com/complianceapi/";

$(document).ready(function () {
  $("#logout").bind("click", function () {
    localStorage.clear();
    window.location = "./taskpane.html";
  });

  $("#file_AttachBtn").click(function () {
    AddAttachment();
  });

  // Company List
  GetMasterCompanylistUserwise();
  // Client Location
  PopulateClientLocationList();
  // Activity-SubActivity
  PopulateActivityListForSearch();

  // Creating Task with Dynamic Values
  $("#newtask_btn").bind("click", function () {
    var newtask = {
      companyID: localStorage.SelectedCompany,
      ClientLocation: $("#ddlClientLocation").val(),
      ComplianceName: $("#emailSubject").val(),
      preparerTaskDuration: $("#preparerTaskDuration").val(),
      reviewerTaskDuration: $("#reviewerTaskDuration").val(),
      preparerTaskDurationType: $("#preparerTaskDurationType").val(),
      reviewerTaskDurationType: $("#reviewerTaskDurationType").val(),
      ActivitSubActivity: $("#ddlActivitysSearch option:selected").text(),
      description: $("#description").val(),
      UserID: localStorage.UserID,
    };
    CreateTask(newtask);
  });
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // For Checking data in newtask.js
    // $("#lblUSerError").html(JSON.stringify(localStorage.SelectedCompany));
    $("#userName").val(Office.context.mailbox.userProfile.emailAddress);
    const item = Office.context.mailbox.item;

    item.body.getAsync("html", processHtmlBody);

    $("#emailSubject").val(item.subject);
    $("#description").val(result.value);
    function processHtmlBody(result) {
      $("#description").val(result.value);
    }
  }
});

// Company List
export async function GetMasterCompanylistUserwise() {
  try {
    var companylist = [];
    var data = {
      storeProcedureName: "GetCompanyandrolebyID",
      parameters: {
        puserID: localStorage.UserID,
      },
    };
    $.ajax({
      type: "POST",
      headers: {
        Authorization: `Bearer ${localStorage.Token}`,
      },
      url: rootURLCompliance + "store/procedure/execute",
      data: JSON.stringify(data),
      dataType: "json",
      async: false,
      contentType: "application/json",
      success: function (resp) {
        localStorage.setItem("SelectedCompany", resp.data.data[0][0].companyfid);

        $("#lblUSerError").html(JSON.stringify(localStorage.SelectedCompany));

        $("#companyselection").empty();
        var companyselection = $("#companyselection");

        $.each(resp.data.data[0], function (data, val) {
          localStorage.setItem("SelectedCompany", resp.data.data[0][0].companyfid);
          var option = $("<option />");
          if (val.roletype.toLowerCase().includes("super")) option.html(val.companyname + " (Super Admin)");
          else option.html(val.companyname + " (" + val.roletype + ")");
          option.val(val.companyfid);
          option.attr("role-id", val.roleId);
          companyselection.append(option);
          companylist.push(val.companyfid);
        });

        $("#companyselection").val(localStorage.SelectedCompany);
        localStorage.setItem("SelectedCompany", resp.data.data[0][0].companyfid);
        //else {
        //    $("#companyselection option:first").attr('selected', 'selected');
        //    localStorage.setItem("SelectedCompany", $('#companyselection').val());
        //}
        //localStorage.setItem("Companylist", JSON.stringify(companylist));
        // $("#companyselection").trigger("chosen:updated");
        // $(".chosen-select1").chosen();
        // if (!($("#companyselection_chosen:contains('Manage Company')").length)) {
        //     $('#companyselection_chosen .chosen-drop').append('<div class="manage_company_wrapper"><a href="/Home/Company" class="addCompany-btn mngComp" >Manage Company(s)</a><a href="/Home/Company" class="addCompany-btn addComp" ><i class="fas fa-plus pr-1"></i> Add Company</a></div>');
        // }
      },
      error: function (xhr, textStatus, errorThrown) {
        $("#item-error").html("<b>GetMasterCompanylistUserwise api Error:</b> <br/>" + errorThrown);
      },
    });
  } catch (error) {
    $("#item-error").html("<b>GetMasterCompanylistUserwise Error:</b> <br/>" + error);
  }
}

// Client Location
export async function PopulateClientLocationList() {
  var data = {
    storeProcedureName: "GetClientAndLocation",
    parameters: {
      userid: localStorage.UserID,
      companyid: localStorage.SelectedCompany,
    },
  };
  $.ajax({
    type: "POST",
    headers: {
      Authorization: `Bearer ${localStorage.Token}`,
    },
    url: rootURLCompliance + "store/procedure/execute",
    data: JSON.stringify(data),
    dataType: "json",
    contentType: "application/json",

    success: function (response, textStatus, xhr) {
      // $("#item-error").html("Data:" + JSON.stringify(response) );

      var filterResult = $.grep(Object(response.data.data[0]), function (j) {
        return j.IsActive == true;
      });
      //console.log(filterResult);

      $("#ddlClientLocation").empty();
      var ddlClientLocation = $("#ddlClientLocation");
      var ddlClientLocation = $("#ddlClientLocation");
      var defaultSelect = $("<option />");
      defaultSelect.html("Others-Registered");
      defaultSelect.val("Others-Registered");
      ddlClientLocation.append(defaultSelect);

      $.each(filterResult, function (index, value) {
        var option = $("<option />");
        option.html(value.Name);
        option.val(value.Name);

        // $("#lblUSerError").html(JSON.stringify(value));

        ddlClientLocation.append(option);
      });

      $("#ddlClientLocation").trigger("chosen:updated");
    },
    error: function (xhr, textStatus, errorThrown) {
      console.log("Error in Operation");
    },
  });
}

// Activity-SubActivity
export async function PopulateActivityListForSearch() {
  var ddlBind1 = "";
  var data = {
    storeProcedureName: "GetActivitiesAndSubActivity",
    parameters: {
      userid: localStorage.UserID,
      companyid: localStorage.SelectedCompany,
    },
  };

  $.ajax({
    type: "POST",
    headers: {
      Authorization: `Bearer ${localStorage.Token}`,
    },
    url: rootURLCompliance + "store/procedure/execute",
    data: JSON.stringify(data),
    dataType: "json",
    contentType: "application/json",

    success: function (response, textStatus, xhr) {
      // debugger;
      var filterResult = $.grep(Object(response.data.data[0]), function (j) {
        return j.IsActive == true;
      });
      $.each(filterResult, function (data, value) {
        ddlBind1 +=
          '<option value="' +
          value.ActivityID +
          "-" +
          value.SubActivityID +
          '" data-id="' +
          value.ActivityID +
          '">' +
          value.Name +
          "</option>";
      });
      //console.log(filterResult);
      $("#ddlActivitysSearch").empty();

      $("#ddlActivitysSearch").append(ddlBind1);

      jQuery("#ddlActivitysSearch").multiselect({
        columns: 1,
        placeholder: "Select Languages",
        enableFiltering: true,
        includeSelectAllOption: true,
      });
    },
    error: function (xhr, textStatus, errorThrown) {
      console.log("Error in Operation");
    },
  });
}

// Create Task
export async function CreateTask(newtask) {
  const item = Office.context.mailbox.item;
  try {
    var form = new FormData();
    form.append("companyID", newtask.companyID);
    form.append("sourceFID", "2"); // Source from email Outlook
    form.append("AccessType", "1");
    form.append("Client", "");
    form.append("Location", "");
    form.append("Activity", "0");
    form.append("SubActivity", "");
    form.append("ComplianceType", "Others");
    form.append("Section", "Others");
    form.append("Priority", "3");
    form.append("ComplianceID", "0");
    form.append("ComplianceName", $("#emailSubject").val());
    form.append("Preparer", "52");
    form.append("Reviewer", "");
    form.append("Activitys", newtask.ActivitSubActivity);
    form.append("ClientLocation", newtask.ClientLocation);
    form.append("Esclation", "");
    form.append("PreparerAlert", "2");
    form.append("ReviewerAlert", "1");
    form.append("EsclationAlert", "2");
    form.append("PreparerAlertDurationType", "2");
    form.append("ReviewerAlertDurationType", "2");
    form.append("EscalationAlertDurationType", "2");
    form.append("preparerTaskDurationType", newtask.preparerTaskDurationType);
    form.append("reviewerTaskDurationType", newtask.reviewerTaskDurationType);
    form.append("preparerDuration", newtask.preparerTaskDuration);
    form.append("reviewerDuration", newtask.reviewerTaskDuration);
    form.append("editor", $("#description").val());
    form.append("Attachments", "[]");
    form.append(
      "SchedulersData",
      '{"frequencyID":"1","frequencyCode":"O","frequencyData":"{\\"OonTime\\":\\"2022-09-26 06:57:00\\"}"}--{"StartDate":"2022-09-26 06:57:00","EndDate":"2022-12-26 06:57:00","NoOfOcurrences":null}'
    );
    form.append("Penalty", "0");
    form.append("UserId", newtask.UserID);
    form.append("CreatedBy", newtask.UserID);
    form.append("CreatedOn", "");
    form.append("complianceTags", "");

    // Using fetch
    var requestOptions = {
      method: "POST",
      body: form,
      redirect: "follow",
    };

    fetch("https://dev.1viewtask.com/ComplianceMaster/CreateUpdateComplience", requestOptions)
      .then((response) => response.text())
      .then((result) => {
        const obj = JSON.parse(result);
        var taskId = obj.taskIDs;

        $.alert({
          type: "success",
          title: "1viewTask Alert",
          content: `Task #${taskId} has been created successfully`,
        });
        //response.render('https://localhost:3000/new_task.html');
        //   window.location.href='https://localhost:3000/new_task.html';
      })
      .catch((error) => {
        $.alert({
          title: "Error",
          content: "400 Bad Request!",
        });
      });
  } catch (error) {
    $.alert({
      title: "Error",
      content: "Something went wrong!!",
    });
  }
}
