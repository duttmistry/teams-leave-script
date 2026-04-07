const puppeteer = require("puppeteer");
const prompt = require("prompt-sync")({ sigint: true });

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Click the first visible interactive element whose text or aria-label matches. */
async function clickByIncludes(page, substrings, options = {}) {
  const { timeout = 15000 } = options;
  const subs = Array.isArray(substrings) ? substrings : [substrings];
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const clicked = await page.evaluate((needles) => {
      const selectors =
        'button, [role="button"], [role="menuitem"], a, [role="tab"], span[role="button"]';
      const nodes = [...document.querySelectorAll(selectors)];
      for (const el of nodes) {
        if (!el.offsetParent) continue;
        const t =
          `${el.innerText || ""} ${el.getAttribute("aria-label") || ""}`.toLowerCase();
        if (needles.every((n) => t.includes(n.toLowerCase()))) {
          el.click();
          return true;
        }
      }
      return false;
    }, subs);
    if (clicked) return true;
    await delay(300);
  }
  return false;
}

async function focusPeoplePicker(page) {
  const handle = await page.evaluateHandle(() => {
    const inputs = [...document.querySelectorAll("input")];
    const pick = (inp) => {
      if (!inp.offsetParent) return false;
      const ph = (inp.getAttribute("placeholder") || "").toLowerCase();
      const al = (inp.getAttribute("aria-label") || "").toLowerCase();
      const blob = `${ph} ${al}`;
      if (blob.includes("group name") || blob.includes("name this group"))
        return false;
      if (inp.getAttribute("type") === "search") return true;
      if (/(name|email|search|people|participant|add|required)/.test(blob))
        return true;
      if (inp.getAttribute("role") === "combobox") return true;
      return false;
    };
    return inputs.find(pick) || null;
  });
  const el = handle.asElement();
  if (el) {
    await el.click({ clickCount: 3 });
    return el;
  }
  return null;
}

async function addEmailsToPicker(page, emails) {
  for (const email of emails) {
    const input = await focusPeoplePicker(page);
    if (!input) {
      throw new Error("Could not find people picker input.");
    }
    await page.keyboard.press("Backspace");
    await page.keyboard.type(email, { delay: 35 });
    await delay(1200);
    await page.keyboard.press("Enter");
    await delay(900);
  }
}

async function fillGroupNameIfPresent(page, name) {
  const handle = await page.evaluateHandle(() => {
    const inputs = [...document.querySelectorAll("input")];
    for (const inp of inputs) {
      if (!inp.offsetParent) continue;
      const ph = (inp.getAttribute("placeholder") || "").toLowerCase();
      const al = (inp.getAttribute("aria-label") || "").toLowerCase();
      if (
        ph.includes("group") ||
        al.includes("group") ||
        ph.includes("chat name") ||
        al.includes("chat name") ||
        (ph.includes("name") && ph.includes("group"))
      ) {
        return inp;
      }
    }
    return null;
  });
  const el = handle.asElement();
  if (el) {
    await el.click({ clickCount: 3 });
    await page.keyboard.press("Backspace");
    await page.keyboard.type(name, { delay: 25 });
  }
}

async function confirmCreateGroup(page) {
  const attempts = [
    ["create", "group"],
    ["new", "group"],
    ["start", "chat"],
    ["start", "group"],
    ["create"],
  ];
  for (const subs of attempts) {
    const ok = await clickByIncludes(page, subs, { timeout: 6000 });
    if (ok) return true;
  }
  return false;
}

async function openNewGroupChatFlow(page) {
  await page.goto("https://teams.microsoft.com/l/chat", {
    waitUntil: "networkidle2",
  });
  await delay(2500);

  let pickedGroup =
    (await clickByIncludes(page, ["new group chat"], { timeout: 10000 })) ||
    (await clickByIncludes(page, ["group chat"], { timeout: 4000 }));

  if (!pickedGroup) {
    const opened =
      (await clickByIncludes(page, ["new chat"], { timeout: 12000 })) ||
      (await clickByIncludes(page, ["compose"], { timeout: 5000 }));
    if (opened) await delay(700);
    pickedGroup =
      (await clickByIncludes(page, ["new group chat"], { timeout: 8000 })) ||
      (await clickByIncludes(page, ["group chat"], { timeout: 5000 }));
  }

  if (!pickedGroup) {
    await clickByIncludes(page, ["group"], { timeout: 3000 });
  }
  await delay(1000);
}

async function autoScrollTeams(page) {
  const panelSelector = '[role="tree"]';

  for (let i = 0; i < 10; i++) {
    await page.evaluate((selector) => {
      const panel = document.querySelector(selector);
      if (panel) panel.scrollTop = panel.scrollHeight;
    }, panelSelector);

    await delay(1000);
  }
}

async function clickMoreOptions(page) {
  return await page.evaluate(() => {
    const btns = [...document.querySelectorAll('button, [role="button"]')];

    const moreBtn = btns.find((b) => {
      const text = (b.innerText || "").toLowerCase();
      const aria = (b.getAttribute("aria-label") || "").toLowerCase();
      return text.includes("more") || aria.includes("more options");
    });

    if (moreBtn) {
      moreBtn.click();
      return true;
    }
    return false;
  });
}

async function clickLeaveIfExists(page) {
  return await page.evaluate(() => {
    const items = [...document.querySelectorAll('[role="menuitem"], button')];

    const leaveBtn = items.find((el) =>
      (el.innerText || "").toLowerCase().includes("leave"),
    );

    if (leaveBtn) {
      leaveBtn.click();
      return true;
    }
    return false;
  });
}

async function confirmLeave(page) {
  return await page.evaluate(() => {
    const btns = [...document.querySelectorAll("button")];

    const confirm = btns.find((b) =>
      (b.innerText || "").toLowerCase().includes("leave"),
    );

    if (confirm) {
      confirm.click();
      return true;
    }
    return false;
  });
}

async function runLeaveTeamsFlow(page) {
  try {
    await page.waitForSelector('[role="treeitem"]', { timeout: 60000 });

    console.log("🔄 Loading all teams...");
    await autoScrollTeams(page);

    const teamElements = await page.$$('[role="treeitem"]');

    console.log(`📊 Total Teams Found: ${teamElements.length}`);

    let success = 0;
    let skipped = 0;
    let failed = 0;

    for (let i = 0; i < teamElements.length; i++) {
      try {
        const team = teamElements[i];
        const name = await page.evaluate((el) => el.innerText, team);

        if (!name.trim()) continue;

        console.log(`\n➡️ Processing: ${name}`);

        // Open team
        await team.click();
        await delay(1500);

        // Click "More options"
        const moreClicked = await clickMoreOptions(page);
        if (!moreClicked) {
          console.log("⚠️ More options not found");
          skipped++;
          continue;
        }

        await delay(1000);

        // Click Leave
        const leaveClicked = await clickLeaveIfExists(page);

        if (!leaveClicked) {
          console.log("⚠️ Cannot leave (owner/restricted)");
          skipped++;
          continue;
        }

        await delay(1000);

        // Confirm Leave
        const confirmed = await confirmLeave(page);

        if (confirmed) {
          console.log("✅ Left successfully");
          success++;
        } else {
          console.log("❌ Confirm not found");
          failed++;
        }

        await delay(2500);
      } catch (err) {
        console.log("❌ Error:", err.message);
        failed++;
      }
    }

    console.log("\n======= FINAL SUMMARY =======");
    console.log(`✅ Left: ${success}`);
    console.log(`⚠️ Skipped: ${skipped}`);
    console.log(`❌ Failed: ${failed}`);
  } catch (err) {
    console.log("🚨 Unexpected error:", err.message);
  }
}

async function runCreateGroupsFlow(page, groupCount, emails) {
  for (let i = 1; i <= groupCount; i++) {
    const groupName = `Testing ${i}`;
    try {
      console.log(`\n[${i}/${groupCount}] Creating "${groupName}"…`);

      await openNewGroupChatFlow(page);

      await page.waitForSelector("input", { timeout: 25000 }).catch(() => {});

      await addEmailsToPicker(page, emails);
      await delay(500);
      await fillGroupNameIfPresent(page, groupName);
      await delay(400);

      const created = await confirmCreateGroup(page);
      if (!created) {
        console.log(
          `Could not find Create/Start for "${groupName}". Check UI layout.`,
        );
      } else {
        console.log(`Submitted "${groupName}".`);
      }

      await delay(5000);
      await page.keyboard.press("Escape").catch(() => {});
      await delay(800);
    } catch (err) {
      console.log(`Error on "${groupName}":`, err.message);
    }
  }
  console.log("\nDone creating groups.");
}

(async () => {
  let browser;
  let page;
  let signedIn = false;

  async function ensureSession() {
    if (!browser) {
      browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
        args: ["--start-maximized"],
      });
      page = await browser.newPage();
    }
    if (!signedIn) {
      console.log("Opening Teams...");
      await page.goto("https://teams.microsoft.com", {
        waitUntil: "networkidle2",
      });
      console.log("Please sign in manually (90 seconds)...");
      await delay(90000);
      signedIn = true;
    }
  }

  let quit = false;

  while (!quit) {
    console.log("\n======== Microsoft Teams helper ========");
    console.log("1 - Leave teams (list & leave)");
    console.log("2 - Create group chats (Testing 1, Testing 2, ...)");
    console.log("0 - Exit");
    const menu = prompt("Choose (0/1/2): ").trim();

    if (menu === "0") {
      quit = true;
      break;
    }

    if (menu === "1") {
      await ensureSession();
      await runLeaveTeamsFlow(page);
    } else if (menu === "2") {
      const countStr = prompt("How many group chats to create? ");
      const groupCount = parseInt(String(countStr).trim(), 10);
      const emailsRaw = prompt("User emails in each group (comma separated): ");
      const emails = String(emailsRaw)
        .split(",")
        .map((e) => e.trim())
        .filter((e) => e.length > 0);

      if (!Number.isFinite(groupCount) || groupCount < 1) {
        console.log("Invalid group count. Use a positive integer.");
        continue;
      }
      if (emails.length === 0) {
        console.log("No emails provided.");
        continue;
      }

      const bad = emails.filter((e) => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
      if (bad.length) {
        console.log("These do not look like valid emails:", bad.join(", "));
        continue;
      }

      const confirmStart = prompt(
        `Create ${groupCount} groups named "Testing 1" … "Testing ${groupCount}" with ${emails.length} member(s) each? (yes/no): `,
      );
      if (String(confirmStart).toLowerCase() !== "yes") {
        console.log("Cancelled.");
        continue;
      }

      await ensureSession();
      await runCreateGroupsFlow(page, groupCount, emails);
    } else {
      console.log("Invalid choice. Enter 0, 1, or 2.");
      continue;
    }

    const back = prompt("Return to main menu? (yes/no): ");
    if (String(back).toLowerCase() !== "yes") {
      quit = true;
    }
  }

  if (browser) {
    await browser.close();
    console.log("Browser closed. Done.");
  } else {
    console.log("Goodbye.");
  }
})();
