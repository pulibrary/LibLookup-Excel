const fs = require('fs')
const path = require('path')
const { JSDOM } = require("jsdom")
const { OfficeMockObject } = require("office-addin-mock");

const mockHostType = {
  Access: 'Access',
  Excel: 'Excel',
  OneNote: 'OneNote',
  Outlook: 'Outlook',
  PowerPoint: 'PowerPoint',
  Project: 'Project',
  Word: 'Word',
};

const mockOfficeInfo = {
  host: mockHostType.Excel,
  platform: "PC" 
};

const mockCells = {
  numberFormat: "",
  values: [],
  format: {
    wrapText: false
  }
}

var mockSection = {startRow: -1}

const mockData = {  
  context: {
    sync: async function () {    
      if(mockSection.startRow >= 0) {        
        for(var i = 0; i < mockSection.values.length; i++) {
          mockCells.values[i+mockSection.startRow] = mockSection.values[i]
        }
        //simulate latency of writing to spreadsheet  
        await jest.advanceTimersByTimeAsync(150)
      }
    },
    officeTheme: {
      isDarkTheme: true
    },
    workbook: {
      worksheets: {
        getActiveWorksheet: function () {
          return this._activeSheet;
        },
        _activeSheet: {
          name: "Grievances",
          getUsedRange: function () {
            return {
              load: function () {}, 
              clear: function () {},             
              getLastRow: function() {
                return {
                  load: function () {},
                  rowIndex: (mockCells.values.length > 0) ? mockCells.values.length-1 : 0
                }
              },
              getLastColumn: function() {
                return {
                  load: function () {},
                  columnIndex: (mockCells.values.length > 0) ? mockCells.values[0].length-1 : 0
                }
              }
                       
            };
          },
          getRangeByIndexes: function(startRow,startColumn,numrows,numcols) {
            const currentLength = mockCells.values.length
            const newLength = startRow + numrows
            if(newLength > currentLength) {
              mockCells.values.push(...Array(newLength - currentLength).fill(['']))
            }
            mockSection = {
              startRow: startRow,
              numberFormat: mockCells.numberFormat,
              values: mockCells.values.slice(startRow,startRow+numrows),
              format: mockCells.format
            }
            return mockSection
          },
          getRange: function() {
            mockSection = {
              load: function () {},
              values: batchRanges,
            }
            return mockSection
          }, 
          getCell: function(row,col) {
            return {
              load: function () {},
              address: String.fromCharCode(65 + col) + row
            }
          }
        }
      },
    }
  },
  run: jest.fn(async (callback) => await callback(mockData.context)),
  onReady: jest.fn((callback) => {
    if (typeof callback === 'function') {
      callback(mockOfficeInfo);
    }
    return Promise.resolve(mockOfficeInfo);
  }),
};

const mockDataError = {
  context: {
    workbook: {
      worksheets: {
        getActiveWorksheet: function () {
          return this._activeSheet;
        },
        _activeSheet: {
          name: "Grievances",
          getUsedRange: function () {
            return {
              clear: function() {},
              getLastRow: function() {
                throw new Error(errorMessage)
              },
              getLastColumn: function() {
                throw new Error(errorMessage)
              }
            }
          },
          getRange: function() {
            throw new Error(errorMessage)
          },
          getRangeByIndexes: function(startRow,startColumn,numrows,numcols) {
            throw new Error(errorMessage)
          }
        }
      }
    },
  },
  run: jest.fn(async (callback) => await callback(mockDataError.context)),
  onReady: jest.fn((callback) => {
    if (typeof callback === 'function') {
      callback(mockOfficeInfo);
    }
    return Promise.resolve(mockOfficeInfo);
  })
}

const errorMessage = "But you can't do that, the critics said!"

const notRunning = "LibLookup Desktop app is not running"
const firstEmptyColumnLetter = "D"
const firstEmptyColumnNumber = 3

const catalogs = ["Grievances","Petitions"]
const query = "date = 1857"
const batchQuery = "id = [[A]] AND date = [[B]]"
const batchRanges = [["1","6"],["2","7"],["3","8"],["4","9"],["5","0"],["",""]]
const fields = ['recno','recno','name','type']
const record = ['Arthur Clenham &amp; co &copy;','Grievance']
const catalogLink = "https://office.circumlocution.go.uk/"
const numberOfRecords = 200
const pageSize = 50

const mockFormHTML = './MockForm.html'
const sidebarHTML = '../taskpane/taskpane.html'
const taskpaneTS = '../taskpane/taskpane.ts'

describe("Excel officejs plugin tests", () => {
  let mockForm;
  let dom;
  
  beforeEach(() => {
    jest.useFakeTimers()
    jest.clearAllMocks()
    mockCells.values = []
    mockSection = {startRow: -1}
    const sidebar = fs.readFileSync(path.resolve(__dirname, sidebarHTML), 'utf8')
    mockForm = fs.readFileSync(path.resolve(__dirname, mockFormHTML ), 'utf8');
    dom = new JSDOM(sidebar)
    global.window = dom.window
    global.document = window.document
    fetchMock = jest.fn()    
    global.fetch = fetchMock
    global.FormData = window.FormData
    global.Event = window.Event
    const officeMock = new OfficeMockObject(mockData)
    global.Office = officeMock
    global.Office.HostType = mockHostType
    mockOfficeInfo.host = mockHostType.Excel
    global.Excel = officeMock
    jest.resetModules()
  })

  afterEach(() => {
    delete global.window;
    delete global.document;
    jest.useRealTimers();
    jest.restoreAllMocks();
  });

  test('Simulate non-Excel launch', async() => {
    mockOfficeInfo.host = "Visual FoxPro 6.0"
    require(taskpaneTS)

    await jest.runAllTimersAsync()
    
    expect(document.querySelector("#content").innerHTML).toBe("")
  })

  test('Simulate LibLookup not running', async () => {  
    fetchMock.mockRejectedValue(new Error("Test LibLookup Not Running"));
    Office.context.officeTheme.isDarkTheme = false
    require(taskpaneTS);

    await jest.runAllTimersAsync()

    expect(document.querySelector("#content").innerHTML).toBe(notRunning)
    expect(document.body.classList).not.toContain("office-dark-mode")

  });

  test('Successfully load form, run in query mode', async () => {
    let batchno = 0
    let mode = "abort"
    
    fetchMock.mockResolvedValue({
      text: async () => mockForm,
      json: async () => { 
        var recordsArray = [fields]
        if(batchno == 0) {
          recordsArray.push(["1","1",...record])
        } else {
          for(i = 1; i <= pageSize; i++) {
            const recno = (batchno-1)*pageSize+i
            recordsArray.push([String(recno),String(recno),...record])
          }
        }
        batchno++
        //simulate fetch latency
        jest.advanceTimersToNextTimer()
        return {
          catalogLink:  (mode == "complete") ? catalogLink : "",
          numberOfRecords: numberOfRecords,
          records: recordsArray
        } 
      }
    });

    const { columnLetterToNumber } = require(taskpaneTS)
    await jest.runAllTimersAsync();

    document.querySelector("select[name='catalog']").innerHTML = 
      catalogs.map(c => `<option>${c}</option>`)
    document.querySelector("select[name='catalog']").value = catalogs[0]
    document.querySelector("input[name='maxRecs']").value = numberOfRecords + 1
    document.querySelector("input[name='q']").value = query
    document.querySelector("input[name='displayFields']").value = fields.join("|")
    document.querySelector('#allRowsButton').dispatchEvent(new window.Event('click'))
    expect(document.querySelector("#endrow").value).toBe(String(1))

    const form = document.querySelector('#queryForm')
    form.dispatchEvent(new Event('submit'), {target: form})
    await jest.advanceTimersToNextTimerAsync()

    expect(document.querySelector("#status").innerHTML).toContain("Retrieved")
    document.querySelector('#abortButton').dispatchEvent(new Event('click'))
    await jest.advanceTimersToNextTimerAsync()

    expect(document.querySelector("#status").innerHTML).toContain("Aborted")
    expect(mockCells.values.length).toBe((pageSize*2)+1) 

    batchno = 0
    mode = "complete"
    form.dispatchEvent(new Event('submit'), {target: form})
    while(document.querySelector("#status").innerHTML != 'Done!') {
      await jest.advanceTimersToNextTimerAsync()
    }
    expect(mockCells.values.length).toBe(numberOfRecords+1)
    expect(document.querySelector('#content').firstElementChild.id).toBe("queryForm")
    const colLetter = document.querySelector('#resultColumn').value
    const colNumber = columnLetterToNumber(colLetter)
    expect(colLetter).toBe(firstEmptyColumnLetter)
    expect(colNumber).toBe(firstEmptyColumnNumber)

    document.querySelector('#startrow').value = String(999)
    document.querySelector('#allRowsButton').dispatchEvent(new window.Event('click'))
    await jest.advanceTimersToNextTimerAsync()
    expect(document.querySelector('#startrow').value).toBe(String(1))
    expect(document.querySelector('#endrow').value).toBe(String(mockCells.values.length))
    expect(document.querySelector("#catalogLink").innerHTML).toContain(catalogLink)
  })

  test('Successfully load form, run in batch mode', async () => {
    let batchno = 0
    fetchMock.mockResolvedValue({
      text: async () => mockForm,
      json: async () => { 
        var recordsArray = [fields]
        for(i = 1; i <= batchno; i++) {
          recordsArray.push([String(batchno),String(batchno),...record])
        }
        batchno++
        //simulate fetch latency
        jest.advanceTimersToNextTimer()
        return {
          catalogLink: "",
          numberOfRecords: recordsArray.length,
          records: recordsArray
        } 
      }
    }); 

    require(taskpaneTS)
    await jest.runAllTimersAsync();
    
    document.querySelector("select[name='catalog']").innerHTML = 
      catalogs.map(c => `<option>${c}</option>`)
    document.querySelector("select[name='catalog']").value = catalogs[1]
    document.querySelector("input[name='maxRecs']").value = numberOfRecords
    document.querySelector("input[name='q']").value = batchQuery
    document.querySelector("input[name='displayFields']").value = fields.join("|")
    document.querySelector("input[name='startrow']").value = 1
    document.querySelector("input[name='endrow']").value = batchRanges[0].length
    document.querySelector("input[name='resultColumn']").value = firstEmptyColumnLetter

    const form = document.querySelector('#queryForm')
    form.dispatchEvent(new Event('submit'), {target: form})
    await jest.advanceTimersToNextTimerAsync()
    expect(document.querySelector("#status").innerHTML).toContain("Searching row")
    while(document.querySelector("#status").innerHTML != 'Done!') {
      await jest.advanceTimersToNextTimerAsync()
    }
    expect(mockCells.values.length).toBe(batchRanges.length)
  })

  test('error catching', async () => {
    results = {records: [], row: 0, column: 0}

    fetchMock.mockResolvedValue({
      text: async () => mockForm,
    }); 

    const officeMockError = new OfficeMockObject(mockDataError)
    global.Office = officeMockError
    global.Office.HostType = mockHostType
    global.Excel = officeMockError

    const {writeToSheet, getRangeValues} = require(taskpaneTS)
    await jest.runAllTimersAsync(); 

    const status = document.querySelector("#status")
    expect(status.innerHTML).toBe(errorMessage)

    status.innerHTML = ""
    await writeToSheet([[]],1,1,false)
    jest.advanceTimersToNextTimerAsync()
    expect(status.innerHTML).toBe(errorMessage)

    status.innerHTML = ""
    await getRangeValues(["A1:A1"])
    jest.advanceTimersToNextTimerAsync()
    expect(status.innerHTML).toBe(errorMessage)


  })   
})