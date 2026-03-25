const puppeteer = require("puppeteer");
const prompt = require("prompt-sync")({ sigint: true });

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    args: ["--start-maximized"],
  });

  const page = await browser.newPage();

  console.log("Opening Teams...");
  await page.goto("https://teams.microsoft.com", { waitUntil: "networkidle2" });

  console.log("Please login manually (60 seconds)...");
  await delay(60000);

  await page.waitForSelector('[role="treeitem"]', { timeout: 60000 });

  const teamElements = await page.$$('[role="treeitem"]');

  const teams = [];

  for (let el of teamElements) {
    const name = await page.evaluate((el) => el.innerText, el);
    if (name.trim()) {
      teams.push({ name, element: el });
    }
  }

  console.log("\n==== Your Teams ====");
  teams.forEach((t, i) => {
    console.log(`${i + 1}. ${t.name}`);
  });

  console.log("\nOptions:");
  console.log("A - Leave ALL teams");
  console.log("Enter team number (example: 3)");
  console.log("Enter multiple numbers (example: 1,4,6)");

  const choice = prompt("Your selection: ");

  let selectedIndexes = [];

  if (choice.toLowerCase() === "a") {
    selectedIndexes = teams.map((_, i) => i);
  } else {
    selectedIndexes = choice
      .split(",")
      .map((num) => parseInt(num.trim()) - 1)
      .filter((i) => i >= 0 && i < teams.length);
  }

  if (selectedIndexes.length === 0) {
    console.log("No valid selection.");
    await browser.close();
    return;
  }

  const confirm = prompt("Are you sure? (yes/no): ");
  if (confirm.toLowerCase() !== "yes") {
    console.log("Cancelled.");
    await browser.close();
    return;
  }

  for (let index of selectedIndexes) {
    try {
      console.log(`Leaving: ${teams[index].name}`);

      // Right click on team
      await teams[index].element.click({ button: "right" });

      // Wait for context menu to appear
      await page.waitForSelector('[role="menu"]', { timeout: 5000 });

      // Find "Leave the team" option (latest Teams text)
      const leaveBtn = await page.evaluateHandle(() => {
        const items = [...document.querySelectorAll('[role="menuitem"]')];
        return items.find((el) => el.innerText.toLowerCase().includes("leave"));
      });

      if (leaveBtn) {
        await leaveBtn.click();
        console.log("Clicked leave option");

        // Wait for confirmation modal
        await page.waitForSelector('[role="dialog"]', { timeout: 5000 });

        // Click confirm leave button
        const confirmBtn = await page.evaluateHandle(() => {
          const buttons = [...document.querySelectorAll("button")];
          return buttons.find((btn) =>
            btn.innerText.toLowerCase().includes("leave"),
          );
        });

        if (confirmBtn) {
          await confirmBtn.click();
          console.log("Left successfully.");
        } else {
          console.log("Confirm button not found.");
        }
      } else {
        console.log("Leave option not found.");
      }

      await delay(4000);
    } catch (err) {
      console.log("Error leaving one team:", err.message);
    }
  }

  console.log("Done.");
})();
