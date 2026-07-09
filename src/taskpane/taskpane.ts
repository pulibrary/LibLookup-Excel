/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var totalResults = -1
const increment = 50
const libLookupDomain = "https://localhost:3950/"

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    try {  
      var abort = false
      const homepageURL = libLookupDomain + "?pageType=plugin"
      fetch(homepageURL).then(response => { 
        response.text().then((form) => {
          document.querySelector('#content').innerHTML = form
          document.querySelector("form").querySelectorAll("script").forEach((script) => {
            var newScript = document.createElement('script')
            newScript.src = script.src
            newScript.type = script.type
            document.head.appendChild(newScript);
          })
          updateStatus("Ready")

          var queryTermField = document.querySelector('#queryTerm')
          document.querySelector('#insertSelectionButton').addEventListener('click', function() {    
            getSelectedRange().then((range) => {
              range = range.replace(/.*!/,'')
              queryTermField.value = `[[${range}]]`
            })      
          })
      
          document.querySelector('#abortButton').addEventListener('click', function() {        
            abort = true
          })

          var queryForm = document.querySelector('#queryForm')
          queryForm?.addEventListener('submit', function(event) {
            event.preventDefault()
            document.getElementById('resultCount').innerHTML = ""
            var formData = new FormData(event.target)
            var catalog = formData.get('catalog')
            var query = formData.get('q')
            var fields = formData.get('displayFields')
            var maxRecs = formData.get('maxRecs')
            var readyForNext = false
            var startRow = 0
            var startColumn = 0
            abort = false
            totalResults = -1    
            document.querySelector('#abortButton').disabled = false
            document.querySelector('#submit').disabled = true
          
            updateStatus("Searching...")

            var queryDecoded = decodeURIComponent(query)
            const batchRanges = queryDecoded.match(/\[\[[^\]]+\]\]/g)?.map(range => range.replace(/^\[\[/,'').replace(/\]\]$/,''))
            var queryBatch = []
          
            if(batchRanges) { //batch mode
              startRow = parseInt(batchRanges.join().match(/\d+/)?.[0])-1
              getRangeValues(batchRanges).then((rangeValues) => {
                for(var i = 0; i < rangeValues[0].length; i++) {
                  var thisQuery = queryDecoded
                  for(var j = 0; j < rangeValues.length; j++) {
                    thisQuery = thisQuery.replace(/\[\[[^\]]+\]\]/,rangeValues[j][i])    
                  }  
                  const queryURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + 
                    encodeURIComponent(thisQuery) + '&displayFields=' + fields + '&start=1&maxRecs=' + maxRecs     
                  queryBatch.push(queryURL)               
                }  
                getFirstEmptyColumn().then((column) => {
                  startColumn = column
                  readyForNext = true
                })
              })                      
            } else { //query mode
              const initialURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + query + '&displayFields=' + fields + '&start=1&maxRecs=1'
              fetch(initialURL).then((data) => {
                data.json().then((tableJSON) => {
                  totalResults = tableJSON.numberOfRecords
                  document.getElementById('resultCount').innerHTML = 'Found ' + totalResults + ' results'
                  var queryMax = Math.min(maxRecs,totalResults)
                  for(var i = 1; i < queryMax; i += increment) {
                    const queryURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + query + '&displayFields=' + fields + '&start=' + i + '&maxRecs=' + Math.min(increment, queryMax-i+1)
                    queryBatch.push(queryURL) 
                  }
                  readyForNext = true  
                })
              })
            }

            var i = 0
            var completeQueries = 0
            var queryInterval = setInterval(() => {
              if(!readyForNext) {
                return
              } 
              if(completeQueries == queryBatch.length || abort) {
                clearInterval(queryInterval)
                if(abort) {
                  updateStatus("Aborted")                  
                } else {
                  updateStatus("Done!")
                }
                document.querySelector('#abortButton').disabled = true
                document.querySelector('#submit').disabled = false
                return
              }
              if(i == queryBatch.length) {
                updateStatus("Finalizing...")
                return
              }   
              readyForNext = false

              const queryURL = queryBatch[i]
              fetch(queryURL).then((data) => {
                data.json().then((tableJSON) => {
                  if(batchRanges) { //batch mode
                    updateStatus("Searching row " + (i+1) + " of " + queryBatch.length)
                    writeToSheet(tableJSON.records.slice(0,2),startRow+i,startColumn,true).then(() => {
                      completeQueries++
                    })           
                  } else { //query mode
                    updateStatus("Retrieved " + Math.min(((i+1)*increment),maxRecs,totalResults) + " of " + Math.min(maxRecs,totalResults) + "  records")                    
                    writeToSheet(tableJSON.records,(i>0) ? (i*increment+1) : 0,0,(i>0)).then(() => {
                      if(i == 0) {
                        i++
                        readyForNext = true  
                      }
                      completeQueries++
                    })            
                  }
                  if(i>0 || batchRanges) {
                    i++
                    readyForNext = true                        
                  }
                })
              })
            },500)    
          })
        })      
      }).catch(error => {
        document.querySelector("#content").innerHTML = "LibLookup Desktop app is not running" 
      })
    } catch(error) {
      console.error(error);
    }
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";    
  }
});

export async function writeToSheet(tableJSON,startRow,startColumn,append) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      if(!append) {
        sheet.getUsedRange().clear();
      } else {
        tableJSON = tableJSON.slice(1)
      } 
      if(!tableJSON || tableJSON.length == 0) {
        return
      }  
      const range = sheet.getRangeByIndexes(startRow,startColumn,tableJSON.length,tableJSON[0].length-1);
      range.numberFormat = [["@"]];
      range.format.wrapText = false
      range.values = tableJSON.map(function(row: Array<string>) { return row.slice(1)});
      await context.sync();
    })
  } catch(error) {
    console.log(error)
  }
}

function updateStatus(status) {
  document.querySelector('#status').innerHTML = status
}

function detectOfficeTheme() {
  // Check if the Office context and theme exist
  if (Office.context && Office.context.officeTheme) {
    var theme = Office.context.officeTheme.diagnostics;
    // theme.background will return "Dark" or "Light"
    if (theme?.background === "Dark") {
      document.body.classList.add("office-dark-mode");
    } else {
        document.body.classList.remove("office-dark-mode");
    }
  }
}

// Run on initialize
Office.onReady(function(info) {
    detectOfficeTheme();
});
/*
Office.context.addHandlerAsync(
    Office.EventType.OfficeThemeChanged,
    detectOfficeTheme()
);
*/

export async function getRangeValues(ranges) {
  var values = []
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        for(var i = 0; i < ranges.length; i++) {
          const selectedRange = sheet.getRange(ranges[i])
          selectedRange.load('values')
          await context.sync()
          values.push(selectedRange.values)
        }
      })
    } catch(error) {
      console.error(error)
    }
  return values
}

export async function getFirstEmptyColumn() {
  var columnNumber = 0
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      const lastColumn = usedRange.getLastColumn();
      lastColumn.load(["columnIndex", "address"]);
      await context.sync();
      columnNumber = lastColumn.columnIndex + 1
    })
  } catch(error) {
    console.error(error)
  }
  return columnNumber
}

export async function getSelectedRange() {
  var rangeAddress = ""
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange()
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const lastRow = sheet.getUsedRange().getLastRow()
      if(range != null) {
        range.load('columnIndex, rowIndex, rowCount, columnCount, address');
        lastRow.load('rowIndex')
        await context.sync()
        var selectedRows = range.rowCount
        if(lastRow.rowIndex < selectedRows) {
          var truncRange = sheet.getRangeByIndexes(0,range.columnIndex,lastRow.rowIndex+1,1)
          truncRange.load('address')
          await context.sync()
          rangeAddress = truncRange.address
        } else {
          rangeAddress = range.address
        }
      }
    })
  } catch(error) {
    console.error(error)
  }
  return rangeAddress
}

export async function runQuery(queryURL) {   
  const xhr = new XMLHttpRequest();
  xhr.open("GET",queryURL, false)
  xhr.send()
  xhr.onload = function() {
    if(xhr.status == 200) {
      const data = JSON.parse(xhr.responseText)    
      return data
    }
  }             
}
