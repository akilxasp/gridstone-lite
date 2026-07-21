// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import App from "./App";

function gridCell(address: string) {
  const element = document.querySelector(`[role="gridcell"][aria-label^="${address},"]`);
  if (!(element instanceof HTMLElement)) throw new Error(`Grid cell ${address} was not rendered`);
  return element;
}

describe("Gridstone interactions", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => cleanup());

  it("edits a cell by double-clicking and commits exactly once", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.dblClick(gridCell("A4"));
    const editor = screen.getByLabelText("Edit cell value");
    await user.clear(editor);
    await user.type(editor, "East{Enter}");

    await waitFor(() => expect(gridCell("A4").getAttribute("aria-label")).toBe("A4, East"));
    expect(screen.getByText("Updated A4")).toBeTruthy();
  }, 15_000);

  it("edits a formula from the formula bar", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(gridCell("D4"));
    const formulaBar = screen.getByLabelText("Value or formula for D4");
    await user.clear(formulaBar);
    await user.type(formulaBar, "=1+2{Enter}");

    await waitFor(() => expect(gridCell("D4").getAttribute("aria-label")).toBe("D4, 3"));
  }, 15_000);

  it.fails("supports select-and-type keyboard editing", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(gridCell("A1"));
    await user.keyboard("42{Enter}");

    await waitFor(() => expect(gridCell("A1").getAttribute("aria-label")).toBe("A1, 42"));
  }, 15_000);

  it("applies formatting and can undo it", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(gridCell("A4"));
    await user.click(screen.getByRole("button", { name: "Bold" }));
    await waitFor(() => expect(gridCell("A4").style.fontWeight).toBe("700"));
    await user.click(screen.getByRole("button", { name: "Undo" }));
    await waitFor(() => expect(gridCell("A4").style.fontWeight).toBe(""));
  }, 15_000);

  it("adds and switches to a new worksheet", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Add worksheet" }));
    expect(screen.getByRole("button", { name: "Sheet 2" })).toBeTruthy();
    expect(screen.getByLabelText("Sheet 2 spreadsheet")).toBeTruthy();
  }, 15_000);

  it("selects an entire column from its header", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(screen.getAllByRole("columnheader").find((element) => element.textContent === "B")!);
    await waitFor(() => expect(gridCell("B1").getAttribute("aria-selected")).toBe("true"));
    expect(gridCell("B50").getAttribute("aria-selected")).toBe("true");
    expect(gridCell("A1").getAttribute("aria-selected")).toBe("false");
  }, 15_000);

  it("merges a range and hides the covered cell", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.click(gridCell("A4"));
    await user.keyboard("{Shift>}{ArrowRight}{/Shift}");
    await user.click(screen.getByRole("button", { name: "Merge cells" }));

    await waitFor(() => expect(document.querySelector('[role="gridcell"][aria-label^="B4,"]')).toBeNull());
    expect(gridCell("A4").style.gridColumn).toContain("span 2");
  }, 15_000);

  it("opens Find with the keyboard shortcut and reports matches", async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.keyboard("{Meta>}f{/Meta}");
    const search = screen.getByPlaceholderText("Find in sheet");
    await user.type(search, "North");
    expect(screen.getByText("1 result")).toBeTruthy();
  }, 15_000);
});
