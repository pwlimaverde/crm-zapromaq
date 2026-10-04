<!-- fonte: https://developer.chrome.com/docs/chromium/headless | obtido em: 2026-09-19 | convertido de HTML -->

- [Home](https://developer.chrome.com/)
- [Docs](https://developer.chrome.com/docs)
- [Automation and testing](https://developer.chrome.com/docs/automation-and-testing)

# Chrome Headless mode Stay organized with collections Save and categorize content based on your preferences.

![Mathias Bynens](https://web.dev/images/authors/mathiasbynens.jpg)

Mathias Bynens

![Peter Kvitek](https://web.dev/images/authors/peterkvitek.jpg)

Peter Kvitek

With Chrome Headless mode, you can run the browser in an unattended environment,
without any visible UI. Essentially, you can run Chrome without chrome.

## Use Headless mode

To use Headless mode, pass the `--headless` command-line flag to a Chrome binary, for example:

### Linux

```
google-chrome --headless
```

### macOS

```
open -a "Google Chrome" --args --headless
```

### Windows

```
start chrome --headless
```

### Use old Headless mode

Previously, Headless mode was
[a separate, alternate browser implementation](https://source.chromium.org/chromium/chromium/src/+/main:headless/;drc=c67febd82ae3e18ac8db1397f4ccfa87b0da2ffc)
that happened to be shipped as part of the same Chrome binary. It didn't share
any of the Chrome browser code in
[`//chrome`](https://source.chromium.org/chromium/chromium/src/+/main:chrome/).

Chrome now has unified Headless and headful modes.

![Headless mode shares code with Chrome.](/static/docs/automation-and-testing/headless/image/the-chrome-headless-is-7ec2038d11f0b.svg)

Since Chrome 132.0.6793.0 the old Headless mode is only available as [a
standalone binary named `chrome-headless-shell`](/blog/chrome-headless-shell) which can be downloaded from [Chrome for Testing dashboard](https://googlechromelabs.github.io/chrome-for-testing/).

### In Puppeteer

To use Headless mode in Puppeteer:

```
import puppeteer from 'puppeteer';

const browser = await puppeteer.launch({
  headless: true,  // (default) enables Chrome Headless mode
  // `headless: 'shell'` enables Headless Shell (old headless)
  // `headless: false` enables "headful" mode
});

const page = await browser.newPage();
await page.goto('https://developer.chrome.com/');

// …

await browser.close();
```

For more information on using Headless in Puppeteer, check out [Headless mode](https://pptr.dev/guides/headless-modes).

### In Selenium-WebDriver

To use Headless mode in Selenium-WebDriver:

```
const driver = await env
  .builder()
  .setChromeOptions(options.addArguments('--headless'))
  .build();

await driver.get('https://developer.chrome.com/');

// …

await driver.quit();
```

See the [Selenium team's blog post](https://www.selenium.dev/blog/2023/headless-is-going-away/#what-are-the-two-headless-modes) for more information, including examples using other language bindings.

## Feedback

We look forward to hearing your feedback about Headless mode. If you
encounter any issues, [file a bug](https://goo.gle/headless-bug).

Except as otherwise noted, the content of this page is licensed under the [Creative Commons Attribution 4.0 License](https://creativecommons.org/licenses/by/4.0/), and code samples are licensed under the [Apache 2.0 License](https://www.apache.org/licenses/LICENSE-2.0). For details, see the [Google Developers Site Policies](https://developers.google.com/site-policies). Java is a registered trademark of Oracle and/or its affiliates.

Last updated 2024-10-21 UTC.

[[["Easy to understand","easyToUnderstand","thumb-up"],["Solved my problem","solvedMyProblem","thumb-up"],["Other","otherUp","thumb-up"]],[["Missing the information I need","missingTheInformationINeed","thumb-down"],["Too complicated / too many steps","tooComplicatedTooManySteps","thumb-down"],["Out of date","outOfDate","thumb-down"],["Samples / code issue","samplesCodeIssue","thumb-down"],["Other","otherDown","thumb-down"]],["Last updated 2024-10-21 UTC."],[],[]]
