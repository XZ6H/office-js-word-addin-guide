/**
 * Unit tests for data binding
 */

import { describe, it, expect } from "vitest";
import {
  generateCustomXml,
  extractDocumentData,
  type ApiBindingConfig,
} from "../src/data-binding";

describe("data-binding", () => {
  describe("generateCustomXml", () => {
    it("should generate valid XML", () => {
      const xml = generateCustomXml(
        "http://example.com/schema",
        "root",
        { name: "John", age: "30" }
      );

      expect(xml).toContain('<?xml version="1.0" encoding="UTF-8"?>');
      expect(xml).toContain('<root xmlns="http://example.com/schema">');
      expect(xml).toContain("<name>John</name>");
      expect(xml).toContain("<age>30</age>");
    });

    it("should escape XML special characters", () => {
      const xml = generateCustomXml(
        "http://test.com",
        "data",
        { content: "Test <script> & more" }
      );

      expect(xml).toContain("&lt;script&gt;");
      expect(xml).toContain("&amp;");
    });
  });

  describe("ApiBindingConfig interface", () => {
    it("should accept valid config", () => {
      const config: ApiBindingConfig = {
        endpoint: "https://api.example.com/data",
        method: "GET",
        headers: { "Authorization": "Bearer token" },
        mapping: { "name-field": "user.name" }
      };

      expect(config.endpoint).toBe("https://api.example.com/data");
      expect(config.mapping["name-field"]).toBe("user.name");
    });
  });
});
