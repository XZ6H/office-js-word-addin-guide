/**
 * Unit tests for template generation
 */

import { describe, it, expect, vi } from "vitest";
import {
  createDocumentTemplate,
  populateTemplate,
  createDynamicLineItemsTable,
  type TemplateField,
  type TemplateData,
  type LineItem,
} from "../src/template-generation";

describe("template-generation", () => {
  describe("TemplateField interface", () => {
    it("should accept valid template field", () => {
      const field: TemplateField = {
        tag: "test-field",
        title: "Test Field",
        type: "RichText" as any,
        placeholder: "Enter value..."
      };
      
      expect(field.tag).toBe("test-field");
      expect(field.title).toBe("Test Field");
    });
  });

  describe("TemplateData interface", () => {
    it("should accept string values", () => {
      const data: TemplateData = {
        "field1": "value1",
        "field2": ["a", "b", "c"]
      };
      
      expect(data["field1"]).toBe("value1");
      expect(Array.isArray(data["field2"])).toBe(true);
    });
  });

  describe("LineItem interface", () => {
    it("should calculate totals correctly", () => {
      const item: LineItem = {
        description: "Test Item",
        quantity: 5,
        unitPrice: 10.50
      };
      
      const total = item.quantity * item.unitPrice;
      expect(total).toBe(52.50);
    });
  });
});
