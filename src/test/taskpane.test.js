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


const mockData = {
  context: {
    officeTheme: {
      isDarkTheme: true
    }
  },
  onReady: jest.fn((callback) => {
    if (typeof callback === 'function') {
      callback(mockOfficeInfo);
    }
    return Promise.resolve(mockOfficeInfo);
  }),
  workbook: {
    range: {
      address: "C2:G3",
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};

const notRunning = "LibLookup Desktop app is not running"
const firstEmptyColumnLetter = "C"
const firstEmptyColumnNumber = 3
const maxRowNumber = 1234

const catalogs = ["Grievances","Petitions"]
const query = "date = 1857"
const batchQuery = "id = [[A]] AND date = [[B]]"
const batchRanges = [[1,2,3,4,5,""],[6,7,8,9,0,""]]
const fields = ['recno','name','type']
const record = ['Arthur Clenham','Grievance']
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
          recordsArray.push([1,1,...record])
        } else {
          for(i = 1; i <= pageSize; i++) {
            const recno = (batchno-1)*pageSize+i
            recordsArray.push([recno,recno,...record])
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

    require(taskpaneTS)
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
    
  })
})