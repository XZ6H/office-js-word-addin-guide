/**
 * Unit tests for find-and-replace operations
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  replaceText,
  replaceInSelection,
  replaceAll,
  replaceAndFormat,
} from "../src/find-and-replace";

// Mock Office.js
vi.mock("@types/office-js", () => ({
  default: {
    run: vi.fn((fn) => fn(mockContext)),
    Document: vi.fn(),
    Body: vi.fn(),
  },
}));

const mockContext = {
  document: {
    body: {
      replace: vi.fn(),
    },
    getSelection: vi.fn(),
  },
  sync: vi.fn(),
};

describe("find-and-replace", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("replaceText", () => {
    it("should return replacement count", async () => {
      const mockResult = { count: 5, load: vi.fn() };
      mockContext.document.body.replace.mockReturnValue(mockResult);

      const count = await replaceText("old", "new");
      
      expect(count).toBe(5);
      expect(mockContext.document.body.replace).toHaveBeenCalledWith("old", "new");
    });

    it("should handle zero replacements", async () => {
      const mockResult = { count: 0, load: vi.fn() };
      mockContext.document.body.replace.mockReturnValue(mockResult);

      const count = await replaceText("nonexistent", "new");
      
      expect(count).toBe(0);
    });
  });

  describe("replaceInSelection", () => {
    it("should replace in selection only", async () => {
      const mockSelection = { replace: vi.fn() };
      const mockResult = { count: 2, load: vi.fn() };
      
      mockContext.document.getSelection.mockReturnValue(mockSelection);
      mockSelection.replace.mockReturnValue(mockResult);

      const count = await replaceInSelection("old", "new");
      
      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(count).toBe(2);
    });
  });

  describe("replaceAll", () => {
    it("should replace all occurrences", async () => {
      const mockResult = { count: 10, load: vi.fn() };
      mockContext.document.body.replace.mockReturnValue(mockResult);

      const count = await replaceAll("old", "new", false);
      
      expect(count).toBe(10);
    });
  });
});
