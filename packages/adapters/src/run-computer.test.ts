import type { ComputerRef } from "@rakazo/adapter-kit";
import { describe, expect, it, vi } from "vitest";
import { isComputerStartFailure, lazyRunComputer } from "./run-computer.js";

describe("lazyRunComputer", () => {
  it("marks the error of a failed start so setup can tell it from other failures", async () => {
    const failure = new TypeError("fetch failed");
    const onFailure = vi.fn();
    const computer = lazyRunComputer(async () => {
      throw failure;
    }, onFailure);

    await expect(computer.get()).rejects.toBe(failure);
    expect(onFailure).toHaveBeenCalledOnce();
    expect(() => computer.throwIfFailed()).toThrow(failure);
    expect(isComputerStartFailure(failure)).toBe(true);
  });

  it("does not mark other errors", async () => {
    const ref: ComputerRef = {
      id: "computer-1",
      botId: "bot-1",
      kind: "docker",
      providerRef: "computer-1",
    };
    const computer = lazyRunComputer(async () => ref);

    await computer.get();
    expect(() => computer.throwIfFailed()).not.toThrow();
    expect(isComputerStartFailure(new TypeError("fetch failed"))).toBe(false);
    expect(isComputerStartFailure("fetch failed")).toBe(false);
    expect(isComputerStartFailure(undefined)).toBe(false);
  });
});
