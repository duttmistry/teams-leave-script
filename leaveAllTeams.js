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

  let continueProcess = true;

  while (continueProcess) {
    try {
      // Reload team list (IMPORTANT after leaving teams)
      await page.waitForSelector('[role="treeitem"]', { timeout: 60000 });

      const teamElements = await page.$$('[role="treeitem"]');
      const teams = [];

      for (let el of teamElements) {
        const name = await page.evaluate((el) => el.innerText, el);
        if (name.trim()) {
          teams.push({ name, element: el });
        }
      }

      if (teams.length === 0) {
        console.log("No teams found.");
        break;
      }

      console.log("\n==== Your Teams ====");
      teams.forEach((t, i) => {
        console.log(`${i + 1}. ${t.name}`);
      });

      console.log("\nOptions:");
      console.log("A - Leave ALL teams");
      console.log("X - Exit");
      console.log("Enter team number (example: 3)");
      console.log("Enter multiple numbers (example: 1,4,6)");

      const choice = prompt("Your selection: ");

      if (choice.toLowerCase() === "x") {
        console.log("Exiting...");
        break;
      }

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
        continue;
      }

      const confirm = prompt("Are you sure? (yes/no): ");
      if (confirm.toLowerCase() !== "yes") {
        console.log("Cancelled.");
        continue;
      }

      // Leave Teams Loop
      for (let index of selectedIndexes) {
        try {
          console.log(`Processing: ${teams[index].name}`);

          await teams[index].element.click({ button: "right" });

          await page.waitForSelector('[role="menu"]', { timeout: 5000 });

          // Get all menu items
          const menuItems = await page.$$(
            '.ms-ContextualMenu-itemText, [role="menuitem"]',
          );

          let actionDone = false;

          for (let item of menuItems) {
            const text = await page.evaluate(
              (el) => el.innerText.toLowerCase(),
              item,
            );

            // 🔴 PRIORITY: DELETE TEAM
            if (text.includes("delete")) {
              console.log("Delete option found → Deleting team");

              await item.click();

              await page.waitForSelector('[role="dialog"]', { timeout: 5000 });

              const confirmBtn = await page.evaluateHandle(() => {
                const buttons = [...document.querySelectorAll("button")];
                return buttons.find((btn) =>
                  btn.innerText.toLowerCase().includes("delete"),
                );
              });

              if (confirmBtn) {
                await confirmBtn.click();
                console.log("Deleted successfully.");
              } else {
                console.log("Delete confirm button not found.");
              }

              actionDone = true;
              break;
            }

            // 🟡 FALLBACK: LEAVE TEAM
            if (text.includes("leave")) {
              console.log("Delete not available → Leaving team");

              await item.click();

              await page.waitForSelector('[role="dialog"]', { timeout: 5000 });

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
                console.log("Leave confirm button not found.");
              }

              actionDone = true;
              break;
            }
          }

          if (!actionDone) {
            console.log("No Delete/Leave option found.");
          }

          await delay(4000);
        } catch (err) {
          console.log("Error processing team:", err.message);
        }
      }

      console.log("\n✅ Operation completed.");

      // Ask user to continue
      const again = prompt("Do you want to continue? (yes/no): ");
      if (again.toLowerCase() !== "yes") {
        continueProcess = false;
      }
    } catch (err) {
      console.log("Unexpected error:", err.message);
      continueProcess = false;
    }
  }

  await browser.close();
  console.log("Browser closed. Done.");
})();
