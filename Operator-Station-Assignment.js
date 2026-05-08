(function () {
  'use strict';

  var tenantRoot = 'https://arlingtonva.sharepoint.com';
  var siteRoot = tenantRoot + '/sites/wpcb';
  var assetsBase = siteRoot + '/forms/SiteAssets/MissingItemsReport/js';
  var jqueryUrl = assetsBase + '/jquery.min.js';
  var jqueryUiUrl = assetsBase + '/jquery-ui.js';
  var spServicesUrl = tenantRoot + '/teams/sys/master/scripts/jquery.SPServices-2014.01.min.js';
  var microsoftAjaxUrl = 'https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js';
  var spJsUrl = tenantRoot + '/_layouts/sp.js';

  function loadScript(url) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.type = 'text/javascript';
      s.src = url;
      s.async = true;
      s.onload = function () { resolve(url); };
      s.onerror = function () { reject(new Error('Failed to load ' + url)); };
      document.head.appendChild(s);
    });
  }

  function loadStyle(cssText) {
    var style = document.createElement('style');
    style.type = 'text/css';
    style.appendChild(document.createTextNode(cssText));
    document.head.appendChild(style);
  }

  // Inject the CSS that was previously inline
  var css = "\
.customtable { color: #333; font-family: Helvetica, Arial, sans-serif; width: 640px; border-collapse: collapse; border-spacing: 0; width:80%; }\
.customtable td, .customtable th { border: 1px solid transparent; height: 30px; transition: all 0.3s; }\
.customtable th { background: #DFDFDF; font-weight: bold; font-size:1.2em; }\
.customtable td { background: #FAFAFA; text-align: left; }\
.customtable tr:nth-child(even) td { background: #F1F1F1; }\
.customtable tr:nth-child(odd) td { background: #FEFEFE; }\
.customtable tr td:hover { background: #666; color: #FFF; }";
  loadStyle(css);

  // Ensure required container exists
  function ensureContainers() {
    var ids = ['divOperatorInfo'];
    ids.forEach(function (id) {
      if (!document.getElementById(id)) {
        var d = document.createElement('div');
        d.id = id;
        document.body.appendChild(d);
      }
    });
  }

  // Table builder
  function makeTable(container, data) {
    var $ = window.jQuery;
    if (!$) {
      console.error('makeTable requires jQuery');
      return;
    }
    var table = $("<table/>").addClass('customtable');
    $.each(data, function (rowIndex, r) {
      var row = $("<tr/>");
      $.each(r, function (colIndex, c) {
        row.append($("<t" + (rowIndex === 0 ? "h" : "d") + "/>").text(c));
      });
      table.append(row);
    });
    return container.append(table);
  }

  // Shift name helper
  function GetShiftNames() {
    var strDate = new Date();
    Date.prototype.yyyymmdd = function () {
      var yyyy = this.getFullYear().toString();
      var mm = (this.getMonth() + 1).toString();
      var dd = this.getDate().toString();
      return yyyy + '_' + (mm[1] ? mm : "0" + mm[0]) + '_' + (dd[1] ? dd : "0" + dd[0]);
    };
    var dateFormatted = strDate.yyyymmdd();
    return [
      'SSA_' + dateFormatted + '_A',
      'SSA_' + dateFormatted + '_B',
      'SSA_' + dateFormatted + '_C',
      'SSA_' + dateFormatted + '_D'
    ];
  }

  // Main data retrieval using SPServices
  function GetOperators() {
    var operatorArray = [];
    var shiftNameArray = GetShiftNames();
    var strCAML = "<Query> <Where> <Or> <Eq> <FieldRef Name='Title' /> <Value Type='Text'>" + shiftNameArray[0] + "</Value> </Eq> <Or> <Eq> <FieldRef Name='Title' /> <Value Type='Text'>" + shiftNameArray[1] + "</Value> </Eq> <Or> <Eq> <FieldRef Name='Title' /> <Value Type='Text'>" + shiftNameArray[2] + "</Value> </Eq> <Eq> <FieldRef Name='Title' /> <Value Type='Text'>" + shiftNameArray[3] + "</Value> </Eq> </Or> </Or> </Or> </Where> </Query>";

    return new Promise(function (resolve, reject) {
      if (!window.jQuery || !window.jQuery.fn || !window.jQuery.fn.SPServices) {
        reject(new Error('jQuery or SPServices not available'));
        return;
      }

      $().SPServices({
        operation: "GetListItems",
        webURL: siteRoot,
        async: false,
        listName: '{C2932C64-DF7B-4221-85D3-5E3A11FDC2D5}',
        CAMLViewFields: "<ViewFields><FieldRef Name='operator' /><FieldRef Name='operatorshift' /><FieldRef Name='Title' /><FieldRef Name='Station' /></ViewFields>",
        CAMLQuery: strCAML,
        CAMLRowLimit: 100,
        completefunc: function (xData, Status) {
          try {
            var responseXML = $.parseXML(xData.responseText);
            var $XML = $(responseXML);
            $XML.find('z\\:row, row').each(function () {
              try {
                var operatorName = $(this).attr('ows_operator');
                var Shift = $(this).attr('ows_operatorshift');
                var Station = $(this).attr('ows_Station');
                operatorArray.push([operatorName, Shift, Station]);
              } catch (err) {
                console.warn('GetOperators row parse error', err);
              }
            });

            if (operatorArray.length > 0) {
              makeTable($('#divOperatorInfo'), operatorArray);
            } else {
              $('#divOperatorInfo').append('There are no items to show in this view of the "Supervisor Station Assignments" list. To add a new item, click "New"');
            }
            resolve(operatorArray);
          } catch (err) {
            reject(err);
          }
        }
      });
    });
  }

  // Entry point
  function ShowOperatorStationAssignment() {
    ensureContainers();
    GetOperators().catch(function (err) {
      console.error('ShowOperatorStationAssignment error', err);
    });
  }

  // Initialization sequence loads dependencies in order
  function init() {
    console.log('missing-items-loader init start');

    // Load jQuery then jQuery UI then SPServices then MicrosoftAjax then sp.js
    loadScript(jqueryUrl)
      .then(function () { return loadScript(jqueryUiUrl); })
      .then(function () { return loadScript(spServicesUrl); })
      .then(function () { return loadScript(microsoftAjaxUrl); })
      .then(function () { return loadScript(spJsUrl); })
      .then(function () {
        console.log('missing-items-loader: all dependencies loaded');
        // small delay to ensure SP scripts are ready
        setTimeout(ShowOperatorStationAssignment, 50);
      })
      .catch(function (err) {
        console.error('missing-items-loader failed to load dependencies', err);
      });
  }

  if (document.readyState === 'complete' || document.readyState === 'interactive') {
    setTimeout(init, 0);
  } else {
    document.addEventListener('DOMContentLoaded', init);
  }

})();
