/**
 * Unit tests for utils.ts
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { safeWordRun, DocumentError } from "../src/utils";

describe("utils", () => {
  describe("DocumentError", () => {
    it("should create error with message", () => {
      const error = new DocumentError("Test error");
      expect(error.message).toBe("Test error");
      expect(error.name).toBe("DocumentError");
    });

    it("should store cause", () => {
      const cause = new Error("Original error");
      const error = new DocumentError("Test error", cause);
      expect(error.cause).toBe(cause);
    });
  });

  describe("safeWordRun", () => {
    it("should return result on success", async () => {
      const mockResult = { text: "test" };
      const fn = vi.fn().mockResolvedValue(mockResult);

      const result = await safeWordRun(fn);
      expect(result).toBe(mockResult);
      expect(fn).toHaveBeenCalledTimes(1);
    });

    it("should throw DocumentError on failure", async () => {
      const fn = vi.fn().mockRejectedValue(new Error("Word error"));

      await expect(safeWordRun(fn)).rejects.toThrow(DocumentError);
    });
  });
});
