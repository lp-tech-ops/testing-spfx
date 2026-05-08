(function () {
  'use strict';

  if (window.__MainPlantDailyLogHistoryLoaded) {
    return;
  }
  window.__MainPlantDailyLogHistoryLoaded = true;

  var hostScript = document.currentScript;

  var tenantRoot = 'https://arlingtonva.sharepoint.com';
  var formsRoot = tenantRoot + '/sites/wpcb/forms';
  var assetsBase = formsRoot + '/SiteAssets/MissingItemsReport/js';

  var jqueryUrl = assetsBase + '/jquery.min.js';
  var jqueryUiUrl = assetsBase + '/jquery-ui.js';
  var spServicesUrl = tenantRoot + '/teams/sys/master/scripts/jquery.SPServices-2014.01.min.js';
  var microsoftAjaxUrl = 'https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js';
  var spJsUrl = tenantRoot + '/_layouts/15/sp.js';
  var spUiDialogUrl = tenantRoot + '/_layouts/15/sp.ui.dialog.js';

  var currlogdate;

  function loadScript(url, testFn) {
    return new Promise(function (resolve, reject) {
      try {
        if (typeof testFn === 'function' && testFn()) {
          resolve(url);
          return;
        }
      } catch (ignore) {}

      var script = document.createElement('script');
      script.type = 'text/javascript';
      script.src = url;
      script.async = true;

      script.onload = function () {
        resolve(url);
      };

      script.onerror = function () {
        reject(new Error('Failed to load script: ' + url));
      };

      document.head.appendChild(script);
    });
  }

  function loadOptionalScript(url, testFn) {
    return loadScript(url, testFn).catch(function (err) {
      console.warn(err.message);
      return null;
    });
  }

  function loadDependencies() {
    return loadScript(jqueryUrl, function () {
      return !!window.jQuery;
    })
      .then(function () {
        return loadScript(jqueryUiUrl, function () {
          return !!(window.jQuery && window.jQuery.ui && window.jQuery.ui.datepicker);
        });
      })
      .then(function () {
        return loadScript(spServicesUrl, function () {
          return !!(window.jQuery && window.jQuery.fn && window.jQuery.fn.SPServices);
        });
      })
      .then(function () {
        return loadOptionalScript(microsoftAjaxUrl, function () {
          return typeof window.Sys !== 'undefined';
        });
      })
      .then(function () {
        return loadOptionalScript(spJsUrl, function () {
          return !!(window.SP && window.SP.ClientContext);
        });
      })
      .then(function () {
        return loadOptionalScript(spUiDialogUrl, function () {
          return !!(window.SP && window.SP.UI && window.SP.UI.ModalDialog);
        });
      });
  }

  function injectCss() {
    if (document.getElementById('mainPlantDailyLogHistoryCss')) {
      return;
    }

    var css =
      ".ui-state-default { color: red; }" +
      ".ui-datepicker-calendar { background-color: white; }" +
      "#mainPlantDailyLogHistoryWp .wpcb-log-toolbar { padding-bottom: 1cm; padding-top: 0.5cm; }" +
      "#mainPlantDailyLogHistoryWp .wpcb-log-datepicker-area { float: right; }" +
      "#mainPlantDailyLogHistoryWp .wpcb-log-link { cursor: pointer; }" +
      "#mainPlantDailyLogHistoryWp .wpcb-toggle-hide, " +
      "#mainPlantDailyLogHistoryWp .wpcb-toggle-show { cursor: pointer; text-decoration: none; }";

    var style = document.createElement('style');
    style.id = 'mainPlantDailyLogHistoryCss';
    style.type = 'text/css';
    style.appendChild(document.createTextNode(css));
    document.head.appendChild(style);
  }

  function renderMarkup() {
    if (document.getElementById('mainPlantDailyLogHistoryWp')) {
      return;
    }

    var html =
      '<div id="mainPlantDailyLogHistoryWp">' +
        '<div align="center" class="ms-rteFontSize-4 ms-rteThemeForeColor-7-0"><u><b>Main Plant Daily Log History</b></u></div>' +
        '<div style="padding-top: 0cm" id="formslogdate" class="ms-rteFontSize-4 ms-rteThemeForeColor-7-0"></div>' +
        '<div class="wpcb-log-toolbar">' +
          '<span><input id="wpcbPrevDayBtn" type="button" value="Previous day" /></span> ' +
          '<span><input id="wpcbTodayBtn" type="button" value="Today" /></span> ' +
          '<span><input id="wpcbNextDayBtn" type="button" value="Next day" /></span> ' +
          '<span class="wpcb-log-datepicker-area">' +
            '<input type="text" id="datepicker" /> ' +
            '<input id="wpcbGetLogsBtn" type="button" value="Get Logs" />' +
          '</span>' +
        '</div>' +
        '<div id="wpcbopdirlog"></div>' +
        '<div id="wpcbprimarylog"></div>' +
        '<div id="wpcbsecondarylog"></div>' +
        '<div id="wpcbawtlog"></div>' +
        '<div id="wpcbdwblog"></div>' +
        '<div id="wpcbliftlog"></div>' +
        '<div id="wpcbsupdirlog"></div>' +
      '</div>';

    if (hostScript && hostScript.parentNode) {
      hostScript.insertAdjacentHTML('beforebegin', html);
    } else {
      document.body.insertAdjacentHTML('beforeend', html);
    }
  }

  function bindEvents() {
    var $ = window.jQuery;

    $('#wpcbPrevDayBtn').off('click.MainPlantDailyLogHistory').on('click.MainPlantDailyLogHistory', prevdaylog);
    $('#wpcbTodayBtn').off('click.MainPlantDailyLogHistory').on('click.MainPlantDailyLogHistory', todaylog);
    $('#wpcbNextDayBtn').off('click.MainPlantDailyLogHistory').on('click.MainPlantDailyLogHistory', nextdaylog);
    $('#wpcbGetLogsBtn').off('click.MainPlantDailyLogHistory').on('click.MainPlantDailyLogHistory', getdaylog);

    $(document)
      .off('click.MainPlantDailyLogHistory', '.wpcb-toggle-hide')
      .on('click.MainPlantDailyLogHistory', '.wpcb-toggle-hide', function (e) {
        e.preventDefault();
        eventshide(this);
      });

    $(document)
      .off('click.MainPlantDailyLogHistory', '.wpcb-toggle-show')
      .on('click.MainPlantDailyLogHistory', '.wpcb-toggle-show', function (e) {
        e.preventDefault();
        eventsshow(this);
      });

    $(document)
      .off('click.MainPlantDailyLogHistory', '.wpcb-log-link')
      .on('click.MainPlantDailyLogHistory', '.wpcb-log-link', function (e) {
        e.preventDefault();
        openLogItemDialog(this.getAttribute('data-url'));
      });
  }

  function pad2(value) {
    value = parseInt(value, 10);
    return value < 10 ? '0' + value : String(value);
  }

  function getmodifieddate(cdate) {
    var dd = pad2(cdate.getDate());
    var mm = pad2(cdate.getMonth() + 1);
    var yyyy = cdate.getFullYear();

    return yyyy + '-' + mm + '-' + dd + 'T12:00:00Z';
  }

  function getcurrenttimewithoffset(offset) {
    var currentdate = new Date();
    currentdate.setDate(currentdate.getDate() - offset);

    var datetime =
      currentdate.getFullYear() + '-' +
      (currentdate.getMonth() + 1) + '-' +
      currentdate.getDate() + 'T' +
      currentdate.getHours() + ':' +
      currentdate.getMinutes() + ':' +
      currentdate.getSeconds();

    return datetime;
  }

  function nextdaylog() {
    if (!currlogdate) {
      currlogdate = new Date();
    }

    currlogdate.setDate(currlogdate.getDate() + 1);
    loadlogdata();
  }

  function prevdaylog() {
    if (!currlogdate) {
      currlogdate = new Date();
    }

    currlogdate.setDate(currlogdate.getDate() - 1);
    loadlogdata();
  }

  function todaylog() {
    currlogdate = new Date();
    loadlogdata();
  }

  function getdaylog() {
    var $ = window.jQuery;
    var selectedDate = null;

    if ($ && $.fn && $.fn.datepicker) {
      selectedDate = $('#datepicker').datepicker('getDate');
    }

    if (!selectedDate) {
      var datepicker = document.getElementById('datepicker');
      var datestring = datepicker ? datepicker.value : '';

      var match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(datestring);

      if (!match) {
        alert('Please enter the date in mm/dd/yyyy format.');
        return;
      }

      var mon1 = parseInt(match[1], 10);
      var dt1 = parseInt(match[2], 10);
      var yr1 = parseInt(match[3], 10);

      selectedDate = new Date(yr1, mon1 - 1, dt1);
    }

    currlogdate = selectedDate;
    loadlogdata();
  }

  function loadlogdata() {
    var $ = window.jQuery;

    if (!currlogdate) {
      currlogdate = new Date();
    }

    var dateLabel = document.getElementById('formslogdate');
    if (dateLabel) {
      dateLabel.innerHTML = 'Log Date: ' + currlogdate.toDateString();
    }

    if ($ && $.fn && $.fn.datepicker && document.getElementById('datepicker')) {
      $('#datepicker').datepicker('setDate', currlogdate);
    }

    var stationQuery =
      '<Query>' +
        '<OrderBy><FieldRef Name="Created"/></OrderBy>' +
        '<Where>' +
          '<Eq>' +
            '<FieldRef Name="Created" />' +
            '<Value Type="DateTime">' + getmodifieddate(currlogdate) + '</Value>' +
          '</Eq>' +
        '</Where>' +
      '</Query>';

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms',
      '{DEC6CAF6-944C-45FA-A14F-9AF71D797452}',
      'Operational Directives Log',
      'wpcbopdirlog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms/primary',
      '{85154071-be7b-4318-9881-be10c8b5f950}',
      'Primary Station Events Log',
      'wpcbprimarylog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms/secondary',
      '{4b77c1a4-ddf3-4f73-8031-8e8ddb1afaf2}',
      'Secondary Station Events Log',
      'wpcbsecondarylog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms/awt',
      '{cd663a51-7318-4ca1-8bc4-f456bf25d094}',
      'AWT Station Log',
      'wpcbawtlog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms/dwb',
      '{b98af892-663a-42d6-b251-fea332f1601d}',
      'Dewater Building Station Events Log',
      'wpcbdwblog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms/lift',
      '{c2f60dcf-486e-4e0c-acab-46354f04a9d1}',
      'Lift Station Events Log',
      'wpcbliftlog',
      stationQuery,
      false
    );

    showwpcblogevents(
      'https://arlingtonva.sharepoint.com/sites/wpcb/forms',
      '{584ebdcd-9498-4fd2-ba47-34ca838a904b}',
      'Supervisors Events Log',
      'wpcbsupdirlog',
      stationQuery,
      false
    );
  }

  function eventshide(e) {
    var $ = window.jQuery;
    var row = $(e).closest('tr');

    row.find('ul.csevents').hide();
    $(e).hide();
    $(e).siblings('.wpcb-toggle-show').show();
  }

  function eventsshow(e) {
    var $ = window.jQuery;
    var row = $(e).closest('tr');

    row.find('ul.csevents').show();
    $(e).hide();
    $(e).siblings('.wpcb-toggle-hide').show();
  }

  function openLogItemDialog(url) {
    if (!url) {
      return;
    }

    if (
      window.SP &&
      window.SP.UI &&
      window.SP.UI.ModalDialog &&
      typeof window.SP.UI.ModalDialog.ShowPopupDialog === 'function'
    ) {
      window.SP.UI.ModalDialog.ShowPopupDialog(url);
    } else {
      window.open(url, '_blank');
    }
  }

  function ensureLogContainer(htmlid) {
    var existing = document.getElementById(htmlid);

    if (existing) {
      return existing;
    }

    var container = document.getElementById('mainPlantDailyLogHistoryWp') || document.body;
    var div = document.createElement('div');
    div.id = htmlid;
    container.appendChild(div);

    return div;
  }

  function htmlEncode(value) {
    if (value === null || typeof value === 'undefined') {
      return '';
    }

    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function attrEncode(value) {
    return htmlEncode(value);
  }

  function valueOrBlank(value) {
    return value === null || typeof value === 'undefined' ? '' : value;
  }

  function showwpcblogevents(siteurl, listname, wptitle, htmlid, myQuery, expand) {
    var $ = window.jQuery;

    if (!$ || !$.fn || !$.fn.SPServices) {
      console.error('SPServices is not available. Cannot load log events for ' + wptitle);
      return;
    }

    ensureLogContainer(htmlid);

    $().SPServices({
      operation: 'GetListItems',
      webURL: siteurl,
      async: false,
      listName: listname,
      CAMLViewFields:
        "<ViewFields>" +
          "<FieldRef Name='KpiDescription' />" +
          "<FieldRef Name='Created' />" +
          "<FieldRef Name='Title' />" +
          "<FieldRef Name='Event_x0020_Category' />" +
        "</ViewFields>",
      CAMLQuery: myQuery,
      CAMLRowLimit: 100,
      completefunc: function (xData, Status) {
        var arrayResults = [];

        try {
          var responseXML = xData.responseXML || $.parseXML(xData.responseText);
          var $XML = $(responseXML);

          $XML.find('z\\:row, row').each(function () {
            try {
              var tcreated = valueOrBlank($(this).attr('ows_Created'));
              var ttitle = valueOrBlank($(this).attr('ows_Title'));
              var tdescription = valueOrBlank($(this).attr('ows_KpiDescription'));
              var teventcategory = valueOrBlank($(this).attr('ows_Event_x0020_Category'));
              var turl = siteurl + '/Lists/eventslog/WPCB%20Event/displayifs.aspx?IsDlg=1&ID=' + $(this).attr('ows_ID');

              arrayResults.push({
                createdcolumn: tcreated,
                titlecolumn: ttitle,
                descriptioncolumn: tdescription,
                linkcolumn: turl,
                eventcategory: teventcategory
              });
            } catch (rowErr) {
              console.warn('Error parsing row for ' + wptitle, rowErr);
            }
          });
        } catch (err) {
          console.error('Error parsing response for ' + wptitle, err);
        }

        arrayResults.sort(function (a, b) {
          var aSort = String(a.titlecolumn || '') + String(a.createdcolumn || '');
          var bSort = String(b.titlecolumn || '') + String(b.createdcolumn || '');

          if (aSort === bSort) {
            return 0;
          }

          return aSort < bSort ? -1 : 1;
        });

        var liHtml =
          "<table width='100%' class='ms-viewlsts wpcb-log-table'>" +
            "<thead>" +
              "<tr style='background-color: #0072c6; color: white; font-weight: bold;'>" +
                "<td height='30px' colspan='2'><b>" + htmlEncode(wptitle) + "</b></td>" +
              "</tr>" +
            "</thead>" +
            "<tbody>";

        if (arrayResults.length === 0) {
          liHtml += "<tr class='ms-alternatingstrong'><td colspan='2' align='center'>No Events</td></tr>";
        } else {
          var currtitle = '';
          var titlecount = 1;

          for (var i = 0; i < arrayResults.length; i++) {
            var item = arrayResults[i];

            if (currtitle === '' || currtitle !== item.titlecolumn) {
              if (currtitle !== '') {
                liHtml += '</ul></td></tr>';
              }

              currtitle = item.titlecolumn;

              var rowOpen;

              if ((titlecount % 2) === 0) {
                rowOpen = "<tr class='ms-alternatingstrong'>";
              } else {
                rowOpen = "<tr style='background-color: white; color: #284775;'>";
              }

              var minusDisplay = expand ? 'block' : 'none';
              var plusDisplay = expand ? 'none' : 'block';
              var eventDisplay = expand ? 'block' : 'none';

              liHtml += rowOpen;
              liHtml +=
                "<td style='width: 30px; vertical-align: top;'>" +
                  "<a href='#' class='wpcb-toggle-hide' style='display:" + minusDisplay + "'><b>-</b></a>" +
                  "<a href='#' class='wpcb-toggle-show' style='display:" + plusDisplay + "'><b>+</b></a>" +
                "</td>";

              liHtml +=
                "<td>" +
                  "<ul>" +
                    "<a href='#' class='wpcb-log-link' data-url='" + attrEncode(item.linkcolumn) + "'>" +
                      "<b><u>" +
                        htmlEncode(item.titlecolumn) +
                        " - <b>[" + htmlEncode(item.eventcategory) + "]</b>" +
                      "</u></b>" +
                    "</a>" +
                  "</ul>" +
                  "<ul style='display:" + eventDisplay + "' class='csevents'>" +
                    "<li>" +
                      item.descriptioncolumn +
                      " <b> - [" + htmlEncode(item.createdcolumn) + "]</b>" +
                    "</li>";

              titlecount++;
            } else {
              liHtml +=
                "<li>" +
                  item.descriptioncolumn +
                  " <b> - [" + htmlEncode(item.createdcolumn) + "]</b>" +
                "</li>";
            }
          }

          liHtml += '</ul></td></tr>';
        }

        liHtml += '</tbody></table>';

        $('#' + htmlid)
          .stop(true, true)
          .hide()
          .html(liHtml)
          .slideDown('slow');
      }
    });
  }

  function init() {
    injectCss();
    renderMarkup();

    loadDependencies()
      .then(function () {
        var $ = window.jQuery;

        $('#pageTitle').hide();

        if ($.fn && $.fn.datepicker) {
          $('#datepicker').datepicker({
            dateFormat: 'mm/dd/yy'
          });
        }

        bindEvents();
        todaylog();
      })
      .catch(function (err) {
        console.error('Main Plant Daily Log History failed to initialize.', err);

        var dateLabel = document.getElementById('formslogdate');
        if (dateLabel) {
          dateLabel.innerHTML =
            'Main Plant Daily Log History failed to load required scripts. Please check the browser console.';
        }
      });
  }

  window.getmodifieddate = getmodifieddate;
  window.getcurrenttimewithoffset = getcurrenttimewithoffset;
  window.nextdaylog = nextdaylog;
  window.prevdaylog = prevdaylog;
  window.todaylog = todaylog;
  window.getdaylog = getdaylog;
  window.loadlogdata = loadlogdata;
  window.eventshide = eventshide;
  window.eventsshow = eventsshow;
  window.showwpcblogevents = showwpcblogevents;

  window.MainPlantDailyLogHistory = {
    nextdaylog: nextdaylog,
    prevdaylog: prevdaylog,
    todaylog: todaylog,
    getdaylog: getdaylog,
    loadlogdata: loadlogdata,
    showwpcblogevents: showwpcblogevents
  };

  if (document.readyState === 'complete' || document.readyState === 'interactive') {
    setTimeout(init, 0);
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }
})();
