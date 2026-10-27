import { test, expect } from "@playwright/test";

test("app loads demo", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByText("Ventas")).toBeVisible();
  await expect(page.getByText("Laptop").first()).toBeVisible();
});

test("formula computes", async ({ page }) => {
  await page.goto("/");
  await page.evaluate(() => (window as any).__store.getState().startEdit(1, 6, "=D2*2"));
  await page.evaluate(() => (window as any).__store.getState().commitEdit());
  const v = await page.evaluate(() => (window as any).__store.getState().engine.getCellValue("Ventas", 1, 6));
  expect(v).toBe(2);
});
