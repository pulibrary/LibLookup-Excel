/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var totalResults = -1
const increment = 50
const libLookupDomain = "https://localhost:3950/"

const entityMap = {
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
  '&apos;': "'",
  '&nbsp;': ' '
};

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
  
          setMaxRowRange()
          document.querySelector('#allRowsButton').addEventListener('click', function() {        
            setMaxRowRange()
          })

          getFirstEmptyColumn().then((column) => {
            document.querySelector("#resultColumn").value = column
          })

          document.querySelector('#abortButton').addEventListener('click', function() {        
            abort = true
          })

          var queryForm = document.querySelector('#queryForm')
          queryForm?.addEventListener('submit', function(event) {
            event.preventDefault()
            document.getElementById('resultCount').innerHTML = ""

            const pollingInterval = 100

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
            var rangeStart = document.querySelector("#startrow").value
            var rangeEnd = document.querySelector("#endrow").value
            const batchRanges = queryDecoded.match(/\[\[[^\]]+\]\]/g)?.map(range => range.replace(/\[\[([A-Z]+)\]\]/g, "$1" + rangeStart + ":$1" + rangeEnd))
            var queryBatch = []
            var catalogLink = ""
          
            if(batchRanges) { //batch mode
              startRow = parseInt(batchRanges.join().match(/\d+/)?.[0])-1
              startColumn = columnLetterToNumber(document.querySelector("#resultColumn").value)
              getRangeValues(batchRanges).then((rangeValues) => {
                for(var i = 0; i < rangeValues[0].length; i++) {
                  var thisQuery = queryDecoded
                  for(var j = 0; j < rangeValues.length; j++) {
                    thisQuery = thisQuery.replace(/\[\[[^\]]+\]\]/,rangeValues[j][i].toString().replaceAll("\"","\\\""))    
                  }  
                  const queryURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + 
                    encodeURIComponent(thisQuery) + '&displayFields=' + fields + '&start=1&maxRecs=' + maxRecs     
                  queryBatch.push(queryURL)               
                }  
                readyForNext = true
              })                      
            } else { //query mode
              const initialURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + encodeURIComponent(queryDecoded) + '&displayFields=' + fields + '&start=1&maxRecs=1'
              fetch(initialURL).then((data) => {
                data.json().then((tableJSON) => {
                  if(tableJSON.catalogLink != "" && catalogLink == "") {
                    catalogLink = tableJSON.catalogLink
                    document.getElementById("catalogLink").innerHTML = `<a target="_blank" href="${catalogLink}">View in Catalog</a>`
                  }
                  totalResults = tableJSON.numberOfRecords
                  document.getElementById('resultCount').innerHTML = 'Found ' + totalResults + ' results'
                  var queryMax = Math.min(maxRecs,totalResults)
                  for(var i = 1; i < queryMax; i += increment) {
                    const queryURL = libLookupDomain + '?format=json&catalog=' + catalog + '&q=' + encodeURIComponent(queryDecoded) + '&displayFields=' + fields + '&start=' + i + '&maxRecs=' + Math.min(increment, queryMax-i+1)
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
                    var combinedRecords = []
                    combinedRecords.push(tableJSON.records[0])
                    var combinedBody = []
                    for(var j = 1; j < tableJSON.records.length; j++) {
                      if(j == 1) {
                        combinedBody = tableJSON.records[1]
                      } else {
                        combinedBody = combinedBody.map((field, index) => field + "|" + tableJSON.records[j][index])
                      }
                    }
                    if(combinedBody.length > 0) {
                      combinedRecords.push(combinedBody)
                    }
                    writeToSheet(combinedRecords,startRow+i,startColumn,true).then(() => {
                      completeQueries++
                    })           
                  } else { //query mode
                    console.log(tableJSON)
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
            },pollingInterval)    
          })
        })      
      }).catch(error => {
        document.querySelector("#content").innerHTML = "LibLookup Desktop app is not running" 
      })
    } catch(error) {
      console.error(error);
    }
  }
});

export async function setMaxRowRange() {
  getMaxUsedRow().then((maxrow) => {
    document.querySelector("#startrow").value = 1
    document.querySelector("#endrow").value = maxrow
  })
}      

export async function writeToSheet(tableJSON,startRow,startColumn,append) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      var numrows = tableJSON.length - 1
      var numcols = tableJSON[0].length - 1

      if(!append) {
        sheet.getUsedRange().clear();
        numrows++
      } else {
        tableJSON = tableJSON.slice(1)
      }  

      if(numrows == 0) {
        tableJSON = [new Array(numcols+1).fill("")]
        numrows = 1
      }

      const range = sheet.getRangeByIndexes(startRow,startColumn,numrows,numcols);
      console.log(startRow,startColumn,numrows,numcols)
      range.numberFormat = [["@"]];
      range.format.wrapText = false
      tableJSON = tableJSON.map(function (row) {return row.map(cell => cell.replace(/&[a-z]+;|&#\d+;/gi, function(match) {
        return entityMap[match] || match;
      }))})
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

Office.onReady(function(info) {
    detectOfficeTheme();
});

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
  var column = ""
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      const lastColumn = usedRange.getLastColumn();
      lastColumn.load("columnIndex");
      await context.sync();
      const firstEmptyColumn = sheet.getCell(1,lastColumn.columnIndex+1)
      
      firstEmptyColumn.load("address")
      await context.sync();          
      console.log(firstEmptyColumn.address)  
      column = firstEmptyColumn.address.replace(/^[^!]+!/,'').replace(/[^A-Z].*/,'')
    })
  } catch(error) {
    console.error(error)
  }
  return column
}

function columnLetterToNumber(letter) {
  let columnNumber = 0;
  let cleanLetter = letter.toUpperCase();
  
  for (let i = 0; i < cleanLetter.length; i++) {
    columnNumber = columnNumber * 26 + (cleanLetter.charCodeAt(i) - 64);
  }
  
  return columnNumber-1;
}

export async function getMaxUsedRow() {
  var lastRowIndex = 1
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const lastRow = sheet.getUsedRange().getLastRow()
      lastRow.load('rowIndex')
      await context.sync()
      lastRowIndex = lastRow.rowIndex+1
    })
  } catch(error) {
    console.error(error)
  }
  return lastRowIndex
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
