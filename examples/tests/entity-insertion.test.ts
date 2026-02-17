/**
 * Unit tests for entity insertion
 */

import { describe, it, expect } from "vitest";
import {
  EntityLibrary,
  entityLibrary,
  insertEntity,
  insertEntityWithFields,
  type StandardEntity,
} from "../src/entity-insertion";

describe("entity-insertion", () => {
  describe("EntityLibrary", () => {
    it("should have default entities loaded", () => {
      const entity = entityLibrary.getEntity("confidentiality-clause");
      expect(entity).toBeDefined();
      expect(entity?.category).toBe("legal");
    });

    it("should filter by category", () => {
      const legalEntities = entityLibrary.getByCategory("legal");
      expect(legalEntities.length).toBeGreaterThan(0);
      expect(legalEntities.every(e => e.category === "legal")).toBe(true);
    });

    it("should return undefined for unknown entity", () => {
      const entity = entityLibrary.getEntity("nonexistent");
      expect(entity).toBeUndefined();
    });
  });

  describe("StandardEntity interface", () => {
    it("should have required properties", () => {
      const entity: StandardEntity = {
        id: "test-entity",
        name: "Test Entity",
        category: "boilerplate",
        content: "Test content",
        version: "1.0",
        lastUpdated: new Date()
      };

      expect(entity.id).toBe("test-entity");
      expect(entity.version).toBe("1.0");
    });
  });
});
