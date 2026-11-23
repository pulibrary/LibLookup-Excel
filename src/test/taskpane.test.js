import { OfficeMockObject } from "office-addin-mock";

const mockData = {
  onReady: jest.fn(() => Promise.resolve({ host: 'Excel', platform: 'PC' })),
  workbook: {
    range: {
      address: "C2:G3",
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};

const officeMock = new OfficeMockObject(mockData);
global.Office = officeMock

const { columnLetterToNumber } = require("../taskpane/taskpane.ts")

describe("Excel tests", () => {
  test("Column letter conversion", async () => {
    expect(columnLetterToNumber("C")).toBe(2)
  })
})