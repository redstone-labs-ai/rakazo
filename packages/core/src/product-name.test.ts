import { describe, expect, it } from "vitest";
import { DEFAULT_PRODUCT_NAME, productNameFromEnv } from "./product-name.js";

describe("productNameFromEnv", () => {
  it("defaults to Rakazo when unset or blank", () => {
    expect(DEFAULT_PRODUCT_NAME).toBe("Rakazo");
    expect(productNameFromEnv({})).toBe("Rakazo");
    expect(productNameFromEnv({ RAKAZO_PRODUCT_NAME: "  " })).toBe("Rakazo");
  });

  it("trims a configured name and refuses one over 40 characters", () => {
    expect(productNameFromEnv({ RAKAZO_PRODUCT_NAME: " Bot " })).toBe("Bot");
    expect(productNameFromEnv({ RAKAZO_PRODUCT_NAME: "x".repeat(40) })).toHaveLength(40);
    expect(() => productNameFromEnv({ RAKAZO_PRODUCT_NAME: "x".repeat(41) })).toThrow(
      "RAKAZO_PRODUCT_NAME must be at most 40 characters",
    );
  });
});
