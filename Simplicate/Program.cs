using System.Collections.ObjectModel;
using System.Text;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using static Scraper;

var pullRequestUrls = new List<string>();

Login();

SelectPullRequestsTab("Active");

SelectMyName();

pullRequestUrls.AddRange(ScrapePullRequestLinks());

SelectPullRequestsTab("Completed");

SelectMyName();

pullRequestUrls.AddRange(ScrapePullRequestLinks());

var pullRequestInfos = pullRequestUrls.Select(ScrapePullRequestInfo).ToList();

// Make a nice overview
var lookup = pullRequestInfos
    .SelectMany(pri => pri.DaysWorkedOn.Select(date => new KeyValuePair<DateTime, PullRequestInfo>(date, pri)))
    .ToLookup(pri => pri.Key, pri => pri.Value);

var sb = new StringBuilder();
foreach (var group in lookup)
{
    sb.AppendLine($"{group.Key:d}:");
    foreach (var pri in group)
    {
        if (pri.WorkItemTitle != null)
        {
            const string regex = "(?:User Story|Bug) ([0-9]{4,6}): (.*)";
            var match = Regex.Match(pri.WorkItemTitle, regex);
            var workItemNr = match.Groups[1].Value;
            sb.Append(workItemNr);
            sb.Append(": ");
            sb.AppendLine($"    - {match.Groups[2].Value}");
        }
        else
        {
            sb.AppendLine($"    - {pri.PullRequestTitle}");
        }

        sb.AppendLine();
    }
}

File.WriteAllText(Path.Combine(SolutionDir, "done.txt"), sb.ToString());

Driver.Quit();

internal static class Scraper
{
    public static IWebDriver Driver = new ChromeDriver();
    public static WebDriverWait Wait = new(Driver, TimeSpan.FromMinutes(5));
    public static string SolutionDir = Directory.GetParent(AppContext.BaseDirectory)!.Parent!.Parent!.Parent!.Parent!.ToString();

    public static PullRequestInfo ScrapePullRequestInfo(string prLink)
    {
        Driver.Navigate().GoToUrl(prLink);

        // Find out date of first commit, ticket nr, ticket title, pr title 
        // let's start with pr title
        var documentTitle = Driver.Title;
        var regex = "Pull request [0-9]+: (.+) - Repos";
        var title = Regex.Match(documentTitle, regex).Groups[1].Value;

        Wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.region-pullRequestDetailsOverviewExtensions")));
        string? workItemTitle = null;
        string? workItemUrl = null;
        try
        {
            var workItemsSection =
                Driver.FindElement(By.CssSelector("div.region-pullRequestDetailsOverviewExtensions"));
            var link = workItemsSection.FindElement(By.CssSelector("a.bolt-link[href]"));
            workItemUrl = link.GetAttribute("href");
            workItemTitle = link.FindElement(By.CssSelector("span.text-ellipsis")).Text;
        }
        catch (NoSuchElementException)
        {
            // no work item linked
        }

        Driver.Navigate().GoToUrl(prLink.TrimEnd('/') + "?_a=commits");
        Wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.repos-commits-table-content")));

        var uniqueCommitHashes = new HashSet<string>();
        var uniqueDates = new HashSet<DateTime>();

        ReadOnlyCollection<IWebElement> theVisibleCommits;

        void processCommits()
        {
            foreach (var visibleCommit in theVisibleCommits)
            {
                // find commit hash
                var secondaryText = visibleCommit.FindElement(By.CssSelector("div.secondary-text"));
                var commitHash = secondaryText.FindElement(By.CssSelector("span.text-ellipsis.monospaced-xs"))
                    .Text;
                var dateLabel = secondaryText
                    .FindElement(By.CssSelector("span.text-ellipsis[aria-label]"))
                    .GetAttribute("aria-label");
                var trimmedDate =
                    Regex.Match(dateLabel, "([0-9]{1,2} [a-z]{2,4} 20[0-9]{2})").Groups[1].Value;
                var parsedDate = DateTime.Parse(
                        trimmedDate.Replace("mrt", "mar").Replace("mei", "may").Replace("okt", "oct")
                    )
                    .Date;
                uniqueDates.Add(parsedDate);
                uniqueCommitHashes.Add(commitHash);
            }
        }

        int newCount, oldCount;
        do
        {
            theVisibleCommits = Driver.FindElements(By.CssSelector("a[role='row']"));
            oldCount = uniqueCommitHashes.Count;
            processCommits();
            newCount = uniqueCommitHashes.Count;
            // scroll the fuck down
            var lastCommit = theVisibleCommits.Last();

            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", lastCommit);
            Thread.Sleep(300);
        } while (newCount > oldCount);

        return new(
            PullRequestTitle: title,
            PullRequestUrl: prLink,
            WorkItemUrl: workItemUrl,
            WorkItemTitle: workItemTitle,
            DaysWorkedOn: uniqueDates
        );
    }

    public static void SelectMyName()
    {
        Driver.FindElement(By.Id("__bolt-identity-picker-downdown-textfield-1")).Click();
        Thread.Sleep(200);
        Driver.FindElement(By.Id("__bolt-identity-picker-downdown-textfield-1")).SendKeys("Erik Jongbloed");

        Wait.Until(
            ExpectedConditions.ElementIsVisible(
                By.CssSelector(
                    "div.bolt-identitypickerdropdown-item.bolt-suggestions-item.bolt-suggestions-isSuggested")));
        var suggestions =
            Driver
                .FindElement(
                    By.CssSelector(
                        "div.bolt-identitypickerdropdown-item.bolt-suggestions-item.bolt-suggestions-isSuggested"))
                .FindElements(By.CssSelector("div.secondary-text"));
        suggestions.Single(element => element.Text.Contains("ejongbloed@digitalrealty.com")).Click();
    }

    public static void SelectPullRequestsTab(string tab)
    {
        Wait.Until(ExpectedConditions.ElementIsVisible(By.Id("__bolt-tab-" + tab.ToLower())));
        Driver.FindElement(By.Id("__bolt-tab-" + tab.ToLower())).Click();
        Wait.Until(
            ExpectedConditions.ElementIsVisible(By.CssSelector("table[aria-label='Pull request table']")));
    }

    public static IEnumerable<string> ScrapePullRequestLinks()
    {
        IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
        Wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table[aria-label='Pull request table'], .vss-ZeroData")));
        if (Driver.FindElements(By.ClassName("vss-ZeroData")).Any())
        {
            return [];
        }
        var pullRequests =
            Driver.FindElements(By.CssSelector("table[aria-label='Pull request table'] a[role='row']"));
        var uniqueLinks = new HashSet<string>();
        foreach (var pr in pullRequests)
        {
            uniqueLinks.Add(pr.GetAttribute("href"));
        }

        // Try to load more 5 times
        for (var i = 0; i < 3; i++)
        {
            var countBefore = uniqueLinks.Count;
            var lastpr = pullRequests.Last();
            // Scroll the last element into view so that it loads more
            js.ExecuteScript("arguments[0].scrollIntoView(true);", lastpr);
            Thread.Sleep(300);
            pullRequests = Driver.FindElements(
                By.CssSelector("table[aria-label='Pull request table'] a[role='row']"));
            // But then the top ones get unloaded so we need to grab whatever is on the screen a couple times and deduplicate
            foreach (var pr in pullRequests)
            {
                uniqueLinks.Add(pr.GetAttribute("href"));
            }

            var countAfter = uniqueLinks.Count;
            if (countBefore == countAfter)
            {
                break;
            }
        }

        return uniqueLinks;
    }

    public static void Login()
    {
        // Load credentials
        var fileLines = File.ReadAllText(Path.Combine(SolutionDir, "creds.txt")).Split(new[]{'\n', '\r'}, StringSplitOptions.RemoveEmptyEntries);
        var username = fileLines[0];
        var password = fileLines[1];
        // Perform login if needed (Azure DevOps may redirect to Microsoft sign-in)
        // Use driver.FindElement to locate and input credentials
        // Create a WebDriverWait to wait for the login page to load

        Driver.Navigate().GoToUrl(@"https://dev.azure.com/inx-buildingmgmt/OpsView/_git/opsview-gui/pullrequest/");
        // Wait until the login button or a specific element is present on the page
        Wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0116")));
        Driver.FindElement(By.Id("i0116")).SendKeys(username);
        Driver.FindElement(By.Id("idSIButton9")).Click();
        Wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0118")));
        Driver.FindElement(By.Id("i0118")).SendKeys(password);
        Wait.Until(ExpectedConditions.ElementIsVisible(By.Id("idSIButton9")));
        Driver.FindElement(By.Id("idSIButton9")).Click();
        // Manual 2FA happens here
        Wait.Until(ExpectedConditions.ElementIsVisible(By.Id("idBtn_Back")));
        Driver.FindElement(By.Id("idBtn_Back")).Click();

    }
}

internal record PullRequestInfo(
    string PullRequestTitle, 
    string PullRequestUrl,
    string? WorkItemUrl, 
    string? WorkItemTitle, 
    HashSet<DateTime> DaysWorkedOn
);
