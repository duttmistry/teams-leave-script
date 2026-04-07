const puppeteer = require("puppeteer");
const prompt = require("prompt-sync")({ sigint: true });

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ✅ Logger
function log(step, msg) {
  console.log(`[${new Date().toISOString()}] [${step}] ${msg}`);
}

/**
 * 🔥 AUTO SCROLL (CORE FIX)
 */
async function autoScrollTeams(page) {
  log("SCROLL", "Scrolling to load all teams...");

  let prevCount = 0;
  let stableCount = 0;

  while (true) {
    const count = await page.$$eval('[role="treeitem"]', (els) => els.length);

    log("SCROLL", `Loaded: ${count}`);

    if (count === prevCount) {
      stableCount++;
    } else {
      stableCount = 0;
    }

    if (stableCount >= 3) {
      log("SCROLL", "End reached.");
      break;
    }

    prevCount = count;

    await page.evaluate(() => {
      const panel =
        document.querySelector('[role="tree"]') ||
        document.querySelector('[data-tid="team-channel-list"]');

      if (panel) {
        panel.scrollBy(0, 1200);
      } else {
        window.scrollBy(0, 1200);
      }
    });

    await delay(1500);
  }
}

/**
 * 🔥 GET ALL TEAMS (FIXED)
 */
async function getAllTeams(page) {
  log("FETCH", "Fetching all teams...");

  await page.waitForSelector('[role="treeitem"]', { timeout: 60000 });

  // 🔥 Scroll fully
  await autoScrollTeams(page);

  // 🔥 Extract unique names
  const names = await page.$$eval('[role="treeitem"]', (els) => {
    const set = new Set();
    const result = [];

    for (let el of els) {
      const name = el.innerText.trim();
      if (!name) continue;

      if (!set.has(name)) {
        set.add(name);
        result.push(name);
      }
    }

    return result;
  });

  log("FETCH", `Total teams found: ${names.length}`);

  // 🔥 Map back to elements
  const elements = await page.$$('[role="treeitem"]');
  const finalTeams = [];

  for (let name of names) {
    for (let el of elements) {
      const text = await page.evaluate((e) => e.innerText.trim(), el);
      if (text === name) {
        finalTeams.push({ name, element: el });
        break;
      }
    }
  }

  return finalTeams;
}

/**
 * SHOW ALL TEAMS
 */
function displayTeams(teams) {
  console.log("\n========= ALL TEAMS =========");
  teams.forEach((t, i) => {
    console.log(`${i + 1}. ${t.name}`);
  });
  console.log("================================\n");
}

/**
 * FILTER LEAVABLE TEAMS
 */
async function filterLeavableTeams(page, teams) {
  log("FILTER", "Checking leave access...");

  const result = [];

  for (let t of teams) {
    try {
      await t.element.click({ button: "right" });
      await page.waitForSelector('[role="menu"]', { timeout: 3000 });

      const items = await page.$$(
        '.ms-ContextualMenu-itemText, [role="menuitem"]',
      );

      let canLeave = false;

      for (let item of items) {
        const text = await page.evaluate(
          (el) => el.innerText.toLowerCase(),
          item,
        );

        if (text.includes("leave")) {
          canLeave = true;
          break;
        }
      }

      if (canLeave) result.push(t);

      await page.keyboard.press("Escape");
      await delay(300);
    } catch (e) {
      log("ERROR", `Check failed: ${t.name}`);
    }
  }

  log("FILTER", `Leavable: ${result.length}`);
  return result;
}

/**
 * LEAVE FLOW
 */
async function runLeaveTeamsFlow(page) {
  log("START", "Leave flow started");

  while (true) {
    const allTeams = await getAllTeams(page);

    if (!allTeams.length) {
      log("INFO", "No teams found");
      return;
    }

    displayTeams(allTeams);

    const teams = await filterLeavableTeams(page, allTeams);

    if (!teams.length) {
      log("INFO", "Nothing to leave");
      return;
    }

    console.log("\n==== LEAVABLE TEAMS ====");
    teams.forEach((t, i) => {
      console.log(`${i + 1}. ${t.name}`);
    });

    console.log("\nA - Leave ALL");
    console.log("X - Exit");

    const choice = prompt("Select: ");

    if (choice.toLowerCase() === "x") break;

    let selected = [];

    if (choice.toLowerCase() === "a") {
      selected = teams.map((_, i) => i);
    } else {
      selected = choice
        .split(",")
        .map((n) => parseInt(n.trim()) - 1)
        .filter((i) => i >= 0 && i < teams.length);
    }

    if (!selected.length) {
      log("WARN", "Invalid input");
      continue;
    }

    const confirm = prompt("Confirm? (yes/no): ");
    if (confirm !== "yes") continue;

    for (let i of selected) {
      const team = teams[i];

      try {
        log("ACTION", `Leaving: ${team.name}`);

        await team.element.click({ button: "right" });
        await page.waitForSelector('[role="menu"]', { timeout: 5000 });

        const items = await page.$$(
          '.ms-ContextualMenu-itemText, [role="menuitem"]',
        );

        for (let item of items) {
          const text = await page.evaluate(
            (el) => el.innerText.toLowerCase(),
            item,
          );

          if (text.includes("leave")) {
            await item.click();

            await page.waitForSelector('[role="dialog"]', {
              timeout: 5000,
            });

            const confirmBtn = await page.evaluateHandle(() => {
              return [...document.querySelectorAll("button")].find((b) =>
                b.innerText.toLowerCase().includes("leave"),
              );
            });

            if (confirmBtn) {
              await confirmBtn.click();
              log("SUCCESS", `Left: ${team.name}`);
            }

            break;
          }
        }

        await delay(3000);
      } catch (e) {
        log("ERROR", `Failed: ${team.name}`);
      }
    }

    log("DONE", "Batch completed");

    const again = prompt("Continue? (yes/no): ");
    if (again !== "yes") break;
  }
}

/**
 * MAIN
 */
(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    args: ["--start-maximized"],
  });

  const page = await browser.newPage();

  log("INIT", "Opening Teams...");
  await page.goto("https://teams.microsoft.com", {
    waitUntil: "networkidle2",
  });

  log("LOGIN", "Login manually (90s)");
  await delay(90000);

  await runLeaveTeamsFlow(page);

  await browser.close();
  log("EXIT", "Done");
})();
