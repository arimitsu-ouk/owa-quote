// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />

(function () {
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();

      $('#insertDefault').click(insertDefault);
      $('#insertMsg1').click(insertMsg1);
      $('#insertMsg2').click(insertMsg2);
      $('#insertMsg3').click(insertMsg3);
      $('#insertCustom').click(insertCustom);
      $('#doQuote').click(doQuote);
    });
  };
  
  function insertText(textToInsert) {
    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text }, 
      function (asyncResult) {
        // Display the result to the user
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
        }
        else {
          app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
        }
      });
  }

  function insertDefault() {
    insertText("Inserted by the Add-in Command Demo add-in.");
  }
  
  function insertMsg1() {
    insertText("Hello World!");
  }
  
  function insertMsg2() {
    insertText("Add-in commands are cool!");
  }
  
  function insertMsg3() {
    insertText("Visit https://dev.outlook.com today for all of your add-in development needs.");
  }
  
  function insertCustom() {
    var textToInsert = $('#textToInsert').val();
    insertText(textToInsert);
  }
  
  function separateBody(bodyLines, separator) {
    var idx = 0;
    for(;idx < bodyLines.length ; idx++) {
      if(bodyLines[idx] === separator) {
        break;
      }
    }
    
    return {
      before : bodyLines.slice(0, idx),
      after : bodyLines.slice(idx)
    };
  }
 
  function insertQuote(quoteLines, bodyLines, insertTop) {
    var body = [];
    
    if(insertTop) {
      // 先頭
      body = quoteLines.concat(bodyLines);
    } else {
      // 末尾
      var sep = "________________________________________";
      var r = separateBody(bodyLines, sep);
      
      body = body.concat(r.before).concat(quoteLines).concat(r.after);
    }
    
    
    
    // 更新
    Office.context.mailbox.item.body.setAsync(
      body.join('\n'),
      { coercionType: Office.CoercionType.Text },
      function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          app.showNotification("Success", "成功しました。多分。");
        } else {
          app.showNotification("Error", "Failed to paste: " + asyncResult.error.message);
        }
      });
  }
  
  function doQuote() {
    var textQuotePrefix = $('#textQuotePrefix').val();
    var insertTop = $("#radioTop").prop('checked');
    
    Office.context.mailbox.item.getSelectedDataAsync(
      Office.CoercionType.Html, 
      this,
      function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          if(asyncResult.value.sourceProperty !== "body" || asyncResult.value.data === "") {
            app.showNotification("Error", "引用部分を選択してください");
            return;
          }
          
          var quoteLines = asyncResult.value.data.split(/\r\n|\r|\n/);
          quoteLines = quoteLines.map(function(value) {
            return textQuotePrefix + value;
          });
          //console.log("quoteLines : %O", quoteLines);
          
          Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Html, 
            {},
            function(asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                var bodyLines = asyncResult.value.split(/\r\n|\r|\n/);
                //console.log("bodyLines : %O", bodyLines);
                insertQuote(quoteLines, bodyLines, insertTop);
              } else {
                app.showNotification("Error", "Failed to quote: " + asyncResult.error.message);
              }
            }
          );          
        }
        else {
          app.showNotification("Error", "Failed to quote: " + asyncResult.error.message);
        }
      }
    );
  }
})();

// MIT License: 
 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 
 
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 
 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.