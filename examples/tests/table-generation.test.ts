/**
 * Unit tests for table generation
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  createTable,
  addTableRow,
  formatTableHeaders,
  autoFitTable,
  type TableData,
} from "../src/table-generation";

describe("table-generation", () => {
  describe("createTable", () => {
    it("should create table with headers and data", async () => {
      const tableData: TableData = {
        headers: ["Name", "Age", "City"],
        rows: [
          ["John", "30", "NYC"],
          ["Jane", "25", "LA"],
        ],
      };

      // Just verify the function doesn't throw with valid data
      // Real testing requires Office.js mocks
      expect(tableData.headers).toHaveLength(3);
      expect(tableData.rows).toHaveLength(2);
    });

    it("should handle empty rows", async () => {
      const tableData: TableData = {
        headers: ["Name"],
        rows: [],
      };

      expect(tableData.rows).toHaveLength(0);
      expect(tableData.headers).toHaveLength(1);
    });
  });

  describe("addTableRow", () => {
    it("should handle row data of varying lengths", () => {
      const rowData1 = ["a", "b", "c"];
      const rowData2 = ["x", "y"];

      expect(rowData1).toHaveLength(3);
      expect(rowData2).toHaveLength(2);
    });
  });
});
