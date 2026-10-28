// ─── MCP Tool definitions & handlers ─────────────────────────────────────────

import type { AzureDevOpsClient, WorkItem } from "./client.ts";
import type { Tool, CallToolResult } from "@modelcontextprotocol/sdk/types.js";

// ═════════════════════════════════════════════════════════════════════════════
// Helpers
// ═════════════════════════════════════════════════════════════════════════════

const ok = (text: string): CallToolResult => ({ content: [{ type: "text", text }] });

function fail(e: unknown): CallToolResult {
  const msg = e instanceof Error ? e.message : String(e);
  // Make common Azure API errors actionable
  let hint = "";
  if (msg.includes("401") || msg.includes("203")) hint = "\n\nHint: credentials may be expired — call auth_status to check.";
  else if (msg.includes("404")) hint = "\n\nHint: resource not found — verify the ID or name exists.";
  else if (msg.includes("No project selected")) hint = "\n\nHint: call switch_project to pick a project first.";
  else if (msg.includes("403")) hint = "\n\nHint: permission denied — check PAT scopes with auth_status.";
  return { content: [{ type: "text", text: `Error: ${msg}${hint}` }], isError: true };
}

function stripHtml(html: string): string {
  return html
    // Convert semantic elements to markdown equivalents
    .replace(/<strong[^>]*>([\s\S]*?)<\/strong>/gi, "**$1**")
    .replace(/<b[^>]*>([\s\S]*?)<\/b>/gi, "**$1**")
    .replace(/<em[^>]*>([\s\S]*?)<\/em>/gi, "_$1_")
    .replace(/<i[^>]*>([\s\S]*?)<\/i>/gi, "_$1_")
    .replace(/<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, "[$2]($1)")
    .replace(/<li[^>]*>/gi, "\n• ")
    .replace(/<\/li>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<p[^>]*>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<h[1-6][^>]*>/gi, "\n### ")
    .replace(/<\/h[1-6]>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    // Decode common HTML entities
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Extract img src URLs from HTML, capped to avoid huge payloads. */
function extractImageUrls(html: string, max = 8): string[] {
  const urls: string[] = [];
  const re = /<img[^>]+src="([^"]+)"/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null && urls.length < max) {
    const src = m[1]!;
    if (src.startsWith("data:")) continue; // already inline, skip
    urls.push(src);
  }
  return urls;
}

// Compact one-line summary used in list views
function fmtWI(wi: WorkItem): string {
  const f = wi.fields;
  const assigned = (f["System.AssignedTo"] as { displayName?: string } | undefined)?.displayName ?? "—";
  const sprint = String(f["System.IterationPath"] ?? "").split("\\").pop() ?? "—";
  const pts = f["Microsoft.VSTS.Scheduling.StoryPoints"];
  const pri = f["Microsoft.VSTS.Common.Priority"];
  const meta = [
    assigned !== "—" ? `@${assigned}` : null,
    sprint !== "—" ? sprint : null,
    pts != null ? `${pts}pts` : null,
    pri != null ? `P${pri}` : null,
  ].filter(Boolean).join(" · ");
  return `#${wi.id} [${f["System.WorkItemType"]}/${f["System.State"]}] ${f["System.Title"]}${meta ? `  — ${meta}` : ""}`;
}

// Full detail view — used by get_work_item
// Rich text fields rendered as their own sections; anything else long or HTML lands under "Other Fields"
const RICH_FIELDS: Record<string, string> = {
  "System.Description": "Description",
  "Microsoft.VSTS.Common.AcceptanceCriteria": "Acceptance Criteria",
  "Microsoft.VSTS.TCM.ReproSteps": "Repro Steps",
  "Microsoft.VSTS.TCM.SystemInfo": "System Info",
};

function fmtWIFull(
  wi: WorkItem,
  comments: Array<{ createdBy: { displayName: string }; createdDate: string; text: string }>,
): string {
  const f = wi.fields;
  const dn = (key: string) =>
    (f[key] as { displayName?: string } | undefined)?.displayName ?? "";

  const section = (title: string, body: string) =>
    body.trim() ? `\n## ${title}\n${body.trim()}` : "";

  // Only include non-empty core fields
  const row = (label: string, val: unknown) =>
    val != null && val !== "" && val !== 0 ? `${label}: ${val}` : "";

  const core = [
    `# #${wi.id} — ${f["System.Title"]}`,
    ``,
    row("Type",    f["System.WorkItemType"]),
    row("State",   f["System.State"]),
    row("Priority", f["Microsoft.VSTS.Common.Priority"]),
    row("Assigned", dn("System.AssignedTo")),
    row("Area",    f["System.AreaPath"]),
    row("Sprint",  f["System.IterationPath"]),
    row("Tags",    f["System.Tags"]),
    row("Points",  f["Microsoft.VSTS.Scheduling.StoryPoints"]),
    row("Remaining", f["Microsoft.VSTS.Scheduling.RemainingWork"] != null ? `${f["Microsoft.VSTS.Scheduling.RemainingWork"]}h` : null),
    row("Estimate",  f["Microsoft.VSTS.Scheduling.OriginalEstimate"] != null ? `${f["Microsoft.VSTS.Scheduling.OriginalEstimate"]}h` : null),
    row("Completed", f["Microsoft.VSTS.Scheduling.CompletedWork"] != null ? `${f["Microsoft.VSTS.Scheduling.CompletedWork"]}h` : null),
    `Created: ${dn("System.CreatedBy")} · ${String(f["System.CreatedDate"] ?? "").slice(0, 10)}`,
    `Updated: ${dn("System.ChangedBy")} · ${String(f["System.ChangedDate"] ?? "").slice(0, 10)}`,
  ].filter(Boolean).join("\n");

  // ── Rich text fields (full text, never truncated) ─────────────────────────
  const rich = (key: string) => (f[key] ? stripHtml(String(f[key])) : "");

  // Custom fields (e.g. Definition of Done) and other long text the core view doesn't cover
  const extra = Object.entries(f)
    .filter(([key, val]) =>
      !(key in RICH_FIELDS) &&
      typeof val === "string" &&
      (key.startsWith("Custom.") || /<\/?[a-z][^>]*>/i.test(val) || val.length > 200),
    )
    .map(([key, val]) => `### ${key.split(".").pop()}\n${stripHtml(String(val))}`)
    .join("\n\n");

  // ── Relations ─────────────────────────────────────────────────────────────
  type Rel = { rel: string; url: string; attributes: Record<string, unknown> };
  const rels: Rel[] = (wi.relations ?? []) as Rel[];

  const wiIdFromUrl = (url: string) => url.split("/").pop() ?? url;

  const parents = rels.filter((r) => r.rel === "System.LinkTypes.Hierarchy-Reverse");
  const children = rels.filter((r) => r.rel === "System.LinkTypes.Hierarchy-Forward");
  const related = rels.filter((r) => r.rel === "System.LinkTypes.Related");
  const deps = rels.filter(
    (r) =>
      r.rel === "System.LinkTypes.Dependency-Forward" ||
      r.rel === "System.LinkTypes.Dependency-Reverse",
  );
  const artifacts = rels.filter((r) => r.rel === "ArtifactLink");
  const prs = artifacts.filter((r) => String(r.attributes["name"] ?? "").toLowerCase().includes("pull request"));
  const commits = artifacts.filter((r) => String(r.attributes["name"] ?? "").toLowerCase().includes("commit"));
  const builds = artifacts.filter((r) => String(r.attributes["name"] ?? "").toLowerCase().includes("build"));

  const fmtArtifact = (r: Rel) => {
    const name = String(r.attributes["name"] ?? "");
    const comment = String(r.attributes["comment"] ?? "");
    // Extract human-readable ID from vstfs URL
    const parts = r.url.replace("vstfs:///", "").split("/");
    const id = parts.at(-1) ?? r.url;
    return `  • ${name}  id:${id}${comment ? `  "${comment}"` : ""}`;
  };

  const relLines: string[] = [];
  if (parents.length) relLines.push(`Parent:   ${parents.map((r) => `#${wiIdFromUrl(r.url)}`).join(", ")}`);
  if (children.length) relLines.push(`Children: ${children.map((r) => `#${wiIdFromUrl(r.url)}`).join(", ")}`);
  if (related.length) relLines.push(`Related:  ${related.map((r) => `#${wiIdFromUrl(r.url)}`).join(", ")}`);
  if (deps.length) {
    relLines.push(
      ...deps.map(
        (r) =>
          `${r.rel === "System.LinkTypes.Dependency-Forward" ? "Successor" : "Predecessor"}: #${wiIdFromUrl(r.url)}`,
      ),
    );
  }

  const prSection = prs.length
    ? `Pull Requests (${prs.length}):\n${prs.map(fmtArtifact).join("\n")}`
    : "";
  const commitSection = commits.length
    ? `Commits (${commits.length}):\n${commits.map(fmtArtifact).join("\n")}`
    : "";
  const buildSection = builds.length
    ? `Builds (${builds.length}):\n${builds.map(fmtArtifact).join("\n")}`
    : "";

  const devLinks = [prSection, commitSection, buildSection].filter(Boolean).join("\n\n");

  // ── Comments (all of them, full text) ─────────────────────────────────────
  const commentBlock =
    comments.length === 0
      ? "No comments."
      : comments
          .map((c) => `[${c.createdDate.slice(0, 10)}] ${c.createdBy.displayName}\n${stripHtml(c.text)}`)
          .join("\n\n");

  return [
    core,
    ...Object.entries(RICH_FIELDS).map(([key, title]) => section(title, rich(key))),
    section("Other Fields", extra),
    relLines.length ? section("Linked Work Items", relLines.join("\n")) : "",
    devLinks ? section("Development & Deployment", devLinks) : "",
    section("Comments", commentBlock),
  ]
    .filter(Boolean)
    .join("\n");
}

/** Hard ceiling on page size so a single call can't flood context. */
const MAX_PAGE = 50;
const DEFAULT_PAGE = 15;

/** Clamp a user-supplied page size into [1, MAX_PAGE]. */
function pageSize(top: number | undefined): number {
  if (top == null) return DEFAULT_PAGE;
  return Math.min(MAX_PAGE, Math.max(1, Math.floor(top)));
}

/**
 * Render a page of work items with a "showing X–Y of N" header and a next-page
 * hint, so the model can fetch more deliberately instead of dumping everything.
 */
function fmtPage(
  items: WorkItem[],
  total: number,
  skip: number,
  capped = false,
): string {
  if (total === 0) return "No results.";
  const start = skip + 1;
  const end = skip + items.length;
  const totalLabel = capped ? `${total}+` : `${total}`;
  const header = `Showing ${start}–${end} of ${totalLabel}`;
  const body = items.map(fmtWI).join("\n\n───\n\n");
  const more =
    end < total
      ? `\n\n— ${capped ? "more" : `${total - end} more`} available. Call again with skip: ${end} for the next page.`
      : "";
  return `${header}:\n\n${body}${more}`;
}

/**
 * Wrap a pre-formatted list body with a "showing X–Y" header and a next-page
 * hint for endpoints that page by offset but don't return a total (commits,
 * PRs). A full page implies more may exist; a short page is the end.
 */
function fmtOffsetPage(body: string, count: number, skip: number, top: number): string {
  const header = `Showing ${skip + 1}–${skip + count}`;
  const more =
    count >= top ? `\n\n— more may be available. Call again with skip: ${skip + count}.` : "";
  return `${header}:\n\n${body}${more}`;
}

const voteLabel = (vote: number) =>
  vote === 10 ? "approved"
  : vote === 5 ? "approved with suggestions"
  : vote === -5 ? "waiting for author"
  : vote === -10 ? "rejected"
  : "no vote";

const indent = (text: string, n: number) => text.replace(/^/gm, " ".repeat(n));

const str = (a: Record<string, unknown>, k: string) => a[k] as string;
const num = (a: Record<string, unknown>, k: string) => a[k] as number;
const opt = <T>(a: Record<string, unknown>, k: string): T | undefined =>
  a[k] != null ? (a[k] as T) : undefined;

// ═════════════════════════════════════════════════════════════════════════════
// Tools
// ═════════════════════════════════════════════════════════════════════════════

const WI_STATES = ["New", "Active", "Resolved", "Closed", "Done", "In Progress", "Removed"];
const WI_TYPES = ["Task", "Bug", "User Story", "Epic", "Feature", "Issue", "Test Case", "Test Plan"];
const PR_STATUSES = ["active", "completed", "abandoned", "all"];
const PR_VOTES: Record<string, number> = {
  approve: 10,
  approve_with_suggestions: 5,
  reset: 0,
  wait_for_author: -5,
  reject: -10,
};
const THREAD_STATUSES = ["active", "fixed", "wontFix", "closed", "byDesign", "pending"];
const MERGE_STRATEGIES = ["squash", "noFastForward", "rebase", "rebaseMerge"];
const BUILD_STATUSES = ["all", "inProgress", "completed", "notStarted"];
const BUILD_RESULTS = ["succeeded", "failed", "canceled", "partiallySucceeded"];
const RELEASE_STATUSES = ["draft", "active", "abandoned"];
const LINK_TYPES = [
  "System.LinkTypes.Hierarchy-Forward",   // parent → child
  "System.LinkTypes.Hierarchy-Reverse",   // child → parent
  "System.LinkTypes.Related",             // related
  "System.LinkTypes.Dependency-Forward",  // successor
  "System.LinkTypes.Dependency-Reverse",  // predecessor
];

export const TOOLS: Tool[] = [
  {
    name: "auth_status",
    description: "Check credentials and list available projects. Call on 401 errors.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "switch_project",
    description: "List all projects or switch the active one. Call first if no project is set.",
    inputSchema: {
      type: "object",
      properties: { project: { type: "string", description: "Project name (omit to list all)" } },
    },
  },
  {
    name: "list_work_items",
    description: "List/search work items. Filters: mine, sprint, state, type, keyword.",
    inputSchema: {
      type: "object",
      properties: {
        mine: { type: "boolean", description: "Only my open items" },
        sprint: { type: "boolean", description: "Only current sprint items" },
        state: { type: "string", enum: WI_STATES },
        type: { type: "string", enum: WI_TYPES },
        keyword: { type: "string", description: "Search title and description" },
        top: { type: "number", description: "Page size (default 15, max 50)" },
        skip: { type: "number", description: "Skip N results for pagination (default 0)" },
      },
    },
  },
  {
    name: "get_work_item",
    description: "Get full detail for one work item — fields, comments, linked PRs/builds.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "number", description: "Work item ID" },
        images: { type: "boolean", description: "Include inline images (slow, large — default false)" },
      },
      required: ["id"],
    },
  },
  {
    name: "create_work_item",
    description: "Create a work item (task, bug, story, epic, etc).",
    inputSchema: {
      type: "object",
      properties: {
        type: { type: "string", enum: WI_TYPES },
        title: { type: "string" },
        description: { type: "string" },
        assignedTo: { type: "string", description: "Email or display name" },
        priority: { type: "number", enum: [1, 2, 3, 4], description: "1=Critical 4=Low" },
        state: { type: "string", enum: WI_STATES },
        iterationPath: { type: "string", description: "Sprint path" },
        areaPath: { type: "string" },
        parentId: { type: "number" },
        storyPoints: { type: "number" },
        tags: { type: "string", description: "Semicolon-separated" },
        acceptanceCriteria: { type: "string" },
        reproSteps: { type: "string", description: "Bug only" },
      },
      required: ["type", "title"],
    },
  },
  {
    name: "update_work_item",
    description: "Update fields on a work item. Only supply fields to change.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "number" },
        title: { type: "string" },
        state: { type: "string", enum: WI_STATES },
        assignedTo: { type: "string", description: "Empty string to unassign" },
        priority: { type: "number", enum: [1, 2, 3, 4] },
        iterationPath: { type: "string" },
        areaPath: { type: "string" },
        storyPoints: { type: "number" },
        tags: { type: "string" },
        comment: { type: "string", description: "Append a discussion comment" },
      },
      required: ["id"],
    },
  },
  {
    name: "add_comment",
    description: "Post a discussion comment on a work item.",
    inputSchema: {
      type: "object",
      properties: {
        workItemId: { type: "number" },
        text: { type: "string" },
      },
      required: ["workItemId", "text"],
    },
  },
  {
    name: "list_comments",
    description: "List discussion comments on a work item, with their IDs (needed to edit/delete).",
    inputSchema: {
      type: "object",
      properties: {
        workItemId: { type: "number" },
      },
      required: ["workItemId"],
    },
  },
  {
    name: "update_comment",
    description: "Edit an existing discussion comment. Get the commentId from list_comments.",
    inputSchema: {
      type: "object",
      properties: {
        workItemId: { type: "number" },
        commentId: { type: "number" },
        text: { type: "string", description: "New comment text (replaces the old)" },
      },
      required: ["workItemId", "commentId", "text"],
    },
  },
  {
    name: "delete_comment",
    description: "Delete a discussion comment. Get the commentId from list_comments.",
    inputSchema: {
      type: "object",
      properties: {
        workItemId: { type: "number" },
        commentId: { type: "number" },
      },
      required: ["workItemId", "commentId"],
    },
  },
  {
    name: "delete_work_item",
    description: "Delete a work item. By default moves to Recycle Bin (recoverable); pass destroy:true to permanently delete.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "number" },
        destroy: { type: "boolean", description: "Permanently destroy instead of recycle bin (default false)" },
      },
      required: ["id"],
    },
  },
  {
    name: "list_child_work_items",
    description: "List the child work items (subtasks) of a work item with state, assignee and points.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "number", description: "Parent work item ID" },
        recursive: { type: "boolean", description: "Walk the whole subtree, not just direct children (default false)" },
      },
      required: ["id"],
    },
  },
  {
    name: "link_work_items",
    description: "Link two work items (parent/child, related, dependency).",
    inputSchema: {
      type: "object",
      properties: {
        sourceId: { type: "number" },
        targetId: { type: "number" },
        linkType: { type: "string", enum: LINK_TYPES },
        comment: { type: "string" },
      },
      required: ["sourceId", "targetId", "linkType"],
    },
  },
  {
    name: "query_wiql",
    description: "Run a raw WIQL query for advanced filtering.",
    inputSchema: {
      type: "object",
      properties: {
        wiql: { type: "string" },
        top: { type: "number", description: "Page size (default 15, max 50)" },
        skip: { type: "number", description: "Skip N results for pagination (default 0)" },
      },
      required: ["wiql"],
    },
  },
  {
    name: "list_repos",
    description: "List all Git repositories in the project.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "list_commits",
    description: "List recent commits in a repository.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        branch: { type: "string" },
        top: { type: "number", description: "Page size (default 20, max 50)" },
        skip: { type: "number", description: "Skip N commits for pagination (default 0)" },
      },
      required: ["repo"],
    },
  },
  {
    name: "get_file",
    description: "Read a file from a repository (up to 500 lines).",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        path: { type: "string", description: "e.g. /src/index.ts" },
        branch: { type: "string" },
      },
      required: ["repo", "path"],
    },
  },
  {
    name: "list_files",
    description: "List files/folders in a repository directory.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        path: { type: "string", description: "Directory path (default: /)" },
        branch: { type: "string" },
      },
      required: ["repo"],
    },
  },
  {
    name: "list_pull_requests",
    description: "List pull requests with status, branch, and reviewer votes.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        status: { type: "string", enum: PR_STATUSES, description: "Default: active" },
        top: { type: "number", description: "Page size (default 20, max 50)" },
        skip: { type: "number", description: "Skip N PRs for pagination (default 0)" },
      },
      required: ["repo"],
    },
  },
  {
    name: "create_pr",
    description: "Create a pull request.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        title: { type: "string" },
        sourceBranch: { type: "string" },
        targetBranch: { type: "string", description: "Default: the repo's default branch" },
        description: { type: "string" },
        isDraft: { type: "boolean" },
        reviewers: { type: "array", items: { type: "string" }, description: "Email addresses" },
        workItemIds: { type: "array", items: { type: "number" } },
        autoComplete: { type: "boolean", description: "Merge (squash, delete source branch) once policies pass" },
      },
      required: ["repo", "title", "sourceBranch"],
    },
  },
  {
    name: "get_pull_request",
    description: "Full detail for one PR: description, branches, merge status, reviewers and votes, linked work items, changed files, open thread count.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
      },
      required: ["repo", "id"],
    },
  },
  {
    name: "list_pr_comments",
    description: "List PR comment threads with file/line context, status, and every reply. Thread IDs are needed to reply or resolve.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        includeResolved: { type: "boolean", description: "Include fixed/closed threads (default true)" },
      },
      required: ["repo", "id"],
    },
  },
  {
    name: "add_pr_comment",
    description: "Comment on a PR. Pass threadId to reply in a thread, filePath (+line) for an inline file comment, or neither for a general comment.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        text: { type: "string", description: "Markdown" },
        threadId: { type: "number", description: "Reply to this thread" },
        filePath: { type: "string", description: "Comment on this file (e.g. /src/app.ts)" },
        line: { type: "number", description: "Line in the new version of the file" },
      },
      required: ["repo", "id", "text"],
    },
  },
  {
    name: "update_pr_thread",
    description: "Set a PR comment thread's status, e.g. resolve it (fixed) or reopen it (active).",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        threadId: { type: "number" },
        status: { type: "string", enum: THREAD_STATUSES },
      },
      required: ["repo", "id", "threadId", "status"],
    },
  },
  {
    name: "vote_pr",
    description: "Vote on a PR as the signed-in user: approve, approve_with_suggestions, wait_for_author, reject, or reset.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        vote: { type: "string", enum: Object.keys(PR_VOTES) },
      },
      required: ["repo", "id", "vote"],
    },
  },
  {
    name: "manage_pr_reviewers",
    description: "Add or remove PR reviewers by email, UPN, or identity ID.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        add: { type: "array", items: { type: "string" } },
        remove: { type: "array", items: { type: "string" } },
        required: { type: "boolean", description: "Mark added reviewers as required (default false)" },
      },
      required: ["repo", "id"],
    },
  },
  {
    name: "update_pr",
    description: "Edit a PR (title, description, draft), set auto-complete, complete (merge), abandon, or reactivate it.",
    inputSchema: {
      type: "object",
      properties: {
        repo: { type: "string" },
        id: { type: "number", description: "Pull request ID" },
        title: { type: "string" },
        description: { type: "string" },
        isDraft: { type: "boolean", description: "false publishes a draft" },
        action: {
          type: "string",
          enum: ["complete", "abandon", "reactivate", "auto_complete", "cancel_auto_complete"],
          description: "complete merges now; auto_complete merges once policies pass",
        },
        mergeStrategy: { type: "string", enum: MERGE_STRATEGIES, description: "For complete/auto_complete (default squash)" },
        deleteSourceBranch: { type: "boolean", description: "For complete/auto_complete (default true)" },
        mergeMessage: { type: "string", description: "Merge commit message for complete/auto_complete" },
      },
      required: ["repo", "id"],
    },
  },
  {
    name: "list_branches",
    description: "List a repo's branches with ahead/behind counts against the default branch.",
    inputSchema: {
      type: "object",
      properties: { repo: { type: "string" } },
      required: ["repo"],
    },
  },
  {
    name: "list_pipelines",
    description: "List all CI/CD pipeline definitions.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "list_builds",
    description: "List recent build runs with status, result, and link.",
    inputSchema: {
      type: "object",
      properties: {
        pipelineId: { type: "number" },
        branch: { type: "string" },
        status: { type: "string", enum: BUILD_STATUSES },
        result: { type: "string", enum: BUILD_RESULTS },
        top: { type: "number", description: "Default 10" },
      },
    },
  },
  {
    name: "run_pipeline",
    description: "Trigger a pipeline run.",
    inputSchema: {
      type: "object",
      properties: {
        pipelineId: { type: "number" },
        branch: { type: "string" },
        variables: { type: "object", additionalProperties: { type: "string" } },
      },
      required: ["pipelineId"],
    },
  },
  {
    name: "get_build_logs",
    description: "Get pipeline/build logs. With no logId or step, returns the full logs and error messages of every failed step, or the list of step logs if nothing failed.",
    inputSchema: {
      type: "object",
      properties: {
        buildId: { type: "number" },
        logId: { type: "number", description: "A specific log ID (from the step list)" },
        step: { type: "string", description: "Step/task/job name (case-insensitive substring), e.g. \"npm test\"" },
        tail: { type: "number", description: "Only the last N lines (default: whole log)" },
      },
      required: ["buildId"],
    },
  },
  {
    name: "get_sprint",
    description: "Get current sprint — name, dates, and team capacity.",
    inputSchema: {
      type: "object",
      properties: { team: { type: "string", description: "Default: first team" } },
    },
  },
  {
    name: "list_sprints",
    description: "List all sprints with start/end dates.",
    inputSchema: {
      type: "object",
      properties: { team: { type: "string", description: "Default: first team" } },
    },
  },
  {
    name: "get_backlog",
    description: "Show the ordered product backlog (priority-ranked work items).",
    inputSchema: {
      type: "object",
      properties: {
        team: { type: "string", description: "Default: first team" },
        level: { type: "string", description: "Backlog level (e.g. Microsoft.RequirementCategory, Microsoft.EpicCategory). Omit for default." },
        top: { type: "number", description: "Page size (default 20, max 50)" },
        skip: { type: "number", description: "Skip N for pagination (default 0)" },
      },
    },
  },
  {
    name: "get_work_item_history",
    description: "Show the change history of a work item — what fields changed, when, and by whom.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "number", description: "Work item ID" },
        top: { type: "number", description: "Most recent N revisions (default 15)" },
      },
      required: ["id"],
    },
  },
  {
    name: "list_team_members",
    description: "List members of a team (display name, email, admin flag).",
    inputSchema: {
      type: "object",
      properties: { team: { type: "string", description: "Default: first team" } },
    },
  },
  {
    name: "list_paths",
    description: "List area or iteration paths — use these exact values for areaPath/iterationPath on create/update.",
    inputSchema: {
      type: "object",
      properties: {
        kind: { type: "string", enum: ["area", "iteration"], description: "Which path tree to list" },
      },
      required: ["kind"],
    },
  },
  {
    name: "get_board",
    description: "Show a team's Kanban board columns and their state mappings.",
    inputSchema: {
      type: "object",
      properties: {
        team: { type: "string", description: "Default: first team" },
        board: { type: "string", description: "Board name/id (omit for first board)" },
      },
    },
  },
  {
    name: "get_build_timeline",
    description: "Show a build's step-by-step timeline — which stage/job/task failed, with error/warning counts.",
    inputSchema: {
      type: "object",
      properties: { buildId: { type: "number" } },
      required: ["buildId"],
    },
  },
  {
    name: "cancel_build",
    description: "Cancel an in-progress build run.",
    inputSchema: {
      type: "object",
      properties: { buildId: { type: "number" } },
      required: ["buildId"],
    },
  },
  {
    name: "list_releases",
    description: "List recent releases with environment deployment status.",
    inputSchema: {
      type: "object",
      properties: {
        definitionId: { type: "number", description: "Filter by release definition" },
        status: { type: "string", enum: RELEASE_STATUSES },
        top: { type: "number", description: "Default 10" },
      },
    },
  },
  {
    name: "create_release",
    description: "Create (and trigger) a release from a release definition.",
    inputSchema: {
      type: "object",
      properties: {
        definitionId: { type: "number" },
        description: { type: "string" },
      },
      required: ["definitionId"],
    },
  },
];

// ═════════════════════════════════════════════════════════════════════════════
// Handler
// ═════════════════════════════════════════════════════════════════════════════

export async function handleTool(
  client: AzureDevOpsClient,
  name: string,
  args: Record<string, unknown>,
): Promise<CallToolResult> {
  try {
    switch (name) {
      // ── Auth ─────────────────────────────────────────────────────────────────
      case "auth_status": {
        if (!client.project) {
          const projects = await client.listProjects();
          return ok(
            [
              `Authenticated — org: ${client.org}`,
              `No project selected yet.`,
              ``,
              `Available projects (${projects.length}):`,
              ...projects.map((p, i) => `  ${i + 1}. ${p.name}${p.description ? `  — ${p.description}` : ""}`),
              ``,
              `Call switch_project(project: "<name>") to select one.`,
            ].join("\n"),
          );
        }
        const p = await client.checkPermissions();
        return ok(
          [
            `Org:     ${client.org}`,
            `Project: ${client.project}`,
            "",
            ...Object.entries(p.details).map(([k, v]) => `  ${v}  ${k}`),
            "",
            p.workItems
              ? "Work items: ok"
              : "Work items: FAILED — check PAT scopes at https://dev.azure.com/{org}/_usersSettings/tokens",
          ].join("\n"),
        );
      }

      case "switch_project": {
        const projects = await client.listProjects();
        const name = opt<string>(args, "project");
        if (!name) {
          return ok(
            [
              `Org: ${client.org}`,
              `Current project: ${client.project ?? "(none)"}`,
              ``,
              `Available projects (${projects.length}):`,
              ...projects.map((p, i) => `  ${i + 1}. ${p.name}${p.description ? `  — ${p.description}` : ""}`),
              ``,
              `Call switch_project(project: "<name>") to switch.`,
            ].join("\n"),
          );
        }
        const match = projects.find(
          (p) => p.name.toLowerCase() === name.toLowerCase(),
        );
        if (!match) {
          return ok(
            `Project "${name}" not found.\n\nAvailable:\n${projects.map((p) => `  • ${p.name}`).join("\n")}`,
          );
        }
        await client.setProject(match.name);
        return ok(`Switched to project: ${match.name}`);
      }

      // ── Work items ───────────────────────────────────────────────────────────
      case "list_work_items": {
        const keyword = opt<string>(args, "keyword");
        const top = pageSize(opt<number>(args, "top"));
        const skip = Math.max(0, opt<number>(args, "skip") ?? 0);
        const wiql = keyword
          ? client.buildSearchWiql({ keyword, state: opt(args, "state"), type: opt(args, "type") })
          : client.buildListWiql({
              assignedToMe: opt<boolean>(args, "mine"),
              currentSprint: opt<boolean>(args, "sprint"),
              state: opt(args, "state"),
              type: opt(args, "type"),
            });
        const { items, total, capped } = await client.pagedWorkItems(wiql, { skip, top });
        return ok(fmtPage(items, total, skip, capped));
      }

      case "get_work_item": {
        const id = num(args, "id");
        const includeImages = opt<boolean>(args, "images") ?? false;
        const [wi, comments] = await Promise.all([
          client.getWorkItem(id),
          client.listComments(id).catch(() => []),
        ]);

        const content: CallToolResult["content"] = [
          { type: "text", text: fmtWIFull(wi, comments) },
        ];

        // Images are opt-in — they're large and slow to download
        if (includeImages) {
          const htmlFields = [
            wi.fields["System.Description"],
            wi.fields["Microsoft.VSTS.Common.AcceptanceCriteria"],
            wi.fields["Microsoft.VSTS.TCM.ReproSteps"],
          ].filter(Boolean).map(String);

          const imgUrls: string[] = [];
          const seen = new Set<string>();
          for (const html of htmlFields) {
            for (const u of extractImageUrls(html, 4)) {
              if (!seen.has(u)) { seen.add(u); imgUrls.push(u); }
              if (imgUrls.length >= 4) break;
            }
          }
          const images = (await Promise.all(imgUrls.map((u) => client.downloadAttachment(u)))).filter(Boolean);
          content.push(...images.map((img) => ({ type: "image" as const, data: img!.data, mimeType: img!.mimeType })));
        }

        return { content };
      }

      case "create_work_item":
        return ok(
          `Created:\n\n${fmtWI(
            await client.createWorkItem({
              type: str(args, "type"),
              title: str(args, "title"),
              description: opt(args, "description"),
              assignedTo: opt(args, "assignedTo"),
              priority: opt(args, "priority"),
              tags: opt(args, "tags"),
              iterationPath: opt(args, "iterationPath"),
              areaPath: opt(args, "areaPath"),
              parentId: opt(args, "parentId"),
              state: opt(args, "state"),
              storyPoints: opt(args, "storyPoints"),
              acceptanceCriteria: opt(args, "acceptanceCriteria"),
              reproSteps: opt(args, "reproSteps"),
            }),
          )}`,
        );

      case "update_work_item":
        return ok(
          `Updated:\n\n${fmtWI(
            await client.updateWorkItem({
              id: num(args, "id"),
              title: opt(args, "title"),
              state: opt(args, "state"),
              assignedTo: opt(args, "assignedTo"),
              priority: opt(args, "priority"),
              iterationPath: opt(args, "iterationPath"),
              areaPath: opt(args, "areaPath"),
              storyPoints: opt(args, "storyPoints"),
              tags: opt(args, "tags"),
              comment: opt(args, "comment"),
            }),
          )}`,
        );

      case "add_comment": {
        const c = await client.addComment(num(args, "workItemId"), str(args, "text"));
        return ok(`Comment #${c.id} added by ${c.createdBy.displayName} at ${c.createdDate.slice(0, 16)}`);
      }

      case "list_comments": {
        const comments = await client.listComments(num(args, "workItemId"));
        if (comments.length === 0) return ok("No comments on this work item.");
        return ok(
          comments
            .map((c) =>
              [
                `#${c.id} — ${c.createdBy.displayName} at ${c.createdDate.slice(0, 16)}`,
                c.text,
              ].join("\n"),
            )
            .join("\n\n"),
        );
      }

      case "update_comment": {
        const c = await client.updateComment(
          num(args, "workItemId"),
          num(args, "commentId"),
          str(args, "text"),
        );
        return ok(`Comment #${c.id} updated.`);
      }

      case "delete_comment": {
        await client.deleteComment(num(args, "workItemId"), num(args, "commentId"));
        return ok(`Comment #${num(args, "commentId")} deleted.`);
      }

      case "delete_work_item": {
        const destroy = opt<boolean>(args, "destroy") ?? false;
        await client.deleteWorkItem(num(args, "id"), destroy);
        return ok(
          destroy
            ? `Work item #${num(args, "id")} permanently destroyed.`
            : `Work item #${num(args, "id")} moved to Recycle Bin (recoverable).`,
        );
      }

      case "list_child_work_items": {
        const rootId = num(args, "id");
        const recursive = opt<boolean>(args, "recursive") ?? false;
        const lines: string[] = [];
        const walk = async (parentId: number, depth: number): Promise<void> => {
          const kids = await client.getWorkItemChildren(parentId);
          for (const k of kids) {
            lines.push(fmtWI(k).replace(/^/gm, "  ".repeat(depth)));
            if (recursive) await walk(k.id, depth + 1);
          }
        };
        await walk(rootId, 0);
        return ok(lines.length ? `Children of #${rootId}:\n\n${lines.join("\n")}` : `#${rootId} has no child work items.`);
      }

      case "link_work_items": {
        const wi = await client.linkWorkItems(
          num(args, "sourceId"),
          num(args, "targetId"),
          str(args, "linkType"),
          opt<string>(args, "comment"),
        );
        const relCount = wi.relations?.length ?? 0;
        return ok(
          `Linked #${num(args, "sourceId")} → #${num(args, "targetId")} (${str(args, "linkType").split(".").pop()})\nWork item #${wi.id} now has ${relCount} relation(s).`,
        );
      }

      case "query_wiql": {
        const top = pageSize(opt<number>(args, "top"));
        const skip = Math.max(0, opt<number>(args, "skip") ?? 0);
        const { items, total, capped } = await client.pagedWorkItems(str(args, "wiql"), { skip, top });
        if (total === 0) return ok("Query returned 0 results.");
        return ok(fmtPage(items, total, skip, capped));
      }

      // ── Builds ───────────────────────────────────────────────────────────────
      case "list_builds": {
        type S = "all" | "inProgress" | "completed" | "cancelling" | "notStarted";
        type R = "succeeded" | "failed" | "canceled" | "partiallySucceeded";
        const builds = await client.listBuilds({
          pipelineId: opt<number>(args, "pipelineId"),
          branch: opt<string>(args, "branch"),
          status: opt<S>(args, "status"),
          result: opt<R>(args, "result"),
          top: opt<number>(args, "top") ?? 10,
        });
        if (builds.length === 0) return ok("No builds found.");
        return ok(
          builds
            .map((b) => {
              const dur =
                b.startTime && b.finishTime
                  ? `${Math.round((new Date(b.finishTime).getTime() - new Date(b.startTime).getTime()) / 60000)}m`
                  : "—";
              const icon =
                b.result === "succeeded" ? "✅"
                : b.result === "failed" ? "❌"
                : b.result === "canceled" ? "⏹"
                : b.status === "inProgress" ? "🔄"
                : "○";
              return [
                `${icon} #${b.id} ${b.buildNumber}`,
                `   Pipeline: ${b.definition.name}`,
                `   Branch:   ${b.sourceBranch.replace("refs/heads/", "")}`,
                `   Duration: ${dur}`,
                `   URL:      ${b._links.web.href}`,
              ].join("\n");
            })
            .join("\n\n"),
        );
      }

      // ── Pull requests ─────────────────────────────────────────────────────────
      case "list_pull_requests": {
        const top = pageSize(opt<number>(args, "top") ?? 20);
        const skip = Math.max(0, opt<number>(args, "skip") ?? 0);
        const prs = await client.listPullRequests(
          str(args, "repo"),
          (opt(args, "status") ?? "active") as "active" | "completed" | "abandoned" | "all",
          top,
          skip,
        );
        if (prs.length === 0) return ok(skip > 0 ? "No more pull requests." : "No pull requests found.");
        const body = prs
          .map((pr) =>
            [
              `#${pr.pullRequestId} [${pr.status}${pr.isDraft ? "/draft" : ""}] ${pr.title}`,
              `  By:     ${pr.createdBy.displayName}`,
              `  Branch: ${pr.sourceRefName.replace("refs/heads/", "")} → ${pr.targetRefName.replace("refs/heads/", "")}`,
              `  Merge:  ${pr.mergeStatus}`,
              pr.reviewers.length
                ? `  Votes:  ${pr.reviewers.map((r) => `${r.displayName}:${voteLabel(r.vote)}`).join(", ")}`
                : "",
            ]
              .filter(Boolean)
              .join("\n"),
          )
          .join("\n\n");
        return ok(fmtOffsetPage(body, prs.length, skip, top));
      }

      case "create_pr": {
        const repo = str(args, "repo");
        const targetBranch =
          opt<string>(args, "targetBranch") ??
          (await client.getRepository(repo)).defaultBranch ??
          "main";
        let pr = await client.createPullRequest({
          repoId: repo,
          title: str(args, "title"),
          sourceBranch: str(args, "sourceBranch"),
          targetBranch,
          description: opt<string>(args, "description"),
          isDraft: opt<boolean>(args, "isDraft"),
          reviewers: opt<string[]>(args, "reviewers"),
          workItemIds: opt<number[]>(args, "workItemIds"),
        });
        if (opt<boolean>(args, "autoComplete")) {
          pr = await client.updatePullRequest(repo, pr.pullRequestId, {
            autoCompleteSetBy: { id: await client.getMyIdentityId() },
            completionOptions: { mergeStrategy: "squash", deleteSourceBranch: true },
          });
        }
        return ok(
          [
            `Pull request created:`,
            `  #${pr.pullRequestId}  ${pr.title}`,
            `  ${pr.sourceRefName.replace("refs/heads/", "")} → ${pr.targetRefName.replace("refs/heads/", "")}`,
            `  Status: ${pr.status}${pr.isDraft ? " (draft)" : ""}`,
            pr.reviewers.length ? `  Reviewers: ${pr.reviewers.map((r) => r.displayName).join(", ")}` : "",
            pr.autoCompleteSetBy ? `  Auto-complete: on` : "",
          ]
            .filter(Boolean)
            .join("\n"),
        );
      }

      case "get_pull_request": {
        const repo = str(args, "repo");
        const id = num(args, "id");
        const [pr, changes, threads, wiIds] = await Promise.all([
          client.getPullRequest(repo, id),
          client.listPullRequestChanges(repo, id).catch(() => []),
          client.listPullRequestThreads(repo, id).catch(() => []),
          client.listPullRequestWorkItems(repo, id).catch(() => []),
        ]);
        const br = (r: string) => r.replace("refs/heads/", "");
        const reviewers = pr.reviewers.map(
          (r) => `  • ${r.displayName}${r.isRequired ? " (required)" : ""}: ${voteLabel(r.vote)}`,
        );
        const userThreads = threads.filter(
          (t) => !t.isDeleted && t.comments.some((c) => c.commentType === "text"),
        );
        const open = userThreads.filter((t) => t.status === "active" || t.status === "pending");
        const files = changes
          .filter((c) => !c.item.isFolder)
          .map((c) => `  ${c.changeType.padEnd(8)} ${c.item.path}${c.originalPath ? `  (from ${c.originalPath})` : ""}`);
        return ok(
          [
            `# PR #${pr.pullRequestId}: ${pr.title}`,
            ``,
            `Status:  ${pr.status}${pr.isDraft ? " (draft)" : ""}`,
            `Branch:  ${br(pr.sourceRefName)} → ${br(pr.targetRefName)}`,
            `Merge:   ${pr.mergeStatus}`,
            `By:      ${pr.createdBy.displayName} · ${pr.creationDate.slice(0, 10)}`,
            pr.closedDate ? `Closed:  ${pr.closedDate.slice(0, 10)}` : "",
            pr.autoCompleteSetBy ? `Auto-complete: set by ${pr.autoCompleteSetBy.displayName}` : "",
            pr.labels?.length ? `Labels:  ${pr.labels.map((l) => l.name).join(", ")}` : "",
            pr.repository?.webUrl ? `URL:     ${pr.repository.webUrl}/pullrequest/${pr.pullRequestId}` : "",
            pr.description ? `\n## Description\n${pr.description}` : "",
            `\n## Reviewers\n${reviewers.length ? reviewers.join("\n") : "  none"}`,
            wiIds.length ? `\n## Linked Work Items\n  ${wiIds.map((w) => `#${w}`).join(", ")}` : "",
            `\n## Comments\n  ${userThreads.length} thread(s), ${open.length} open. Use list_pr_comments to read them.`,
            `\n## Changed Files (${files.length})\n${files.join("\n") || "  none"}`,
          ]
            .filter(Boolean)
            .join("\n"),
        );
      }

      case "list_pr_comments": {
        const includeResolved = opt<boolean>(args, "includeResolved") ?? true;
        const threads = (await client.listPullRequestThreads(str(args, "repo"), num(args, "id")))
          .filter((t) => !t.isDeleted && t.comments.some((c) => c.commentType === "text"))
          .filter((t) => includeResolved || t.status === "active" || t.status === "pending");
        if (threads.length === 0) return ok("No comment threads.");
        return ok(
          threads
            .map((t) => {
              const ctx = t.threadContext;
              const line = ctx?.rightFileStart?.line ?? ctx?.leftFileStart?.line;
              const where = ctx ? `${ctx.filePath}${line ? `:${line}` : ""}` : "general";
              const comments = t.comments
                .filter((c) => c.commentType === "text" && !c.isDeleted)
                .map((c) => `  ${c.parentCommentId ? "↳ " : ""}${c.author.displayName} · ${c.publishedDate.slice(0, 16).replace("T", " ")}\n${indent(c.content ?? "", c.parentCommentId ? 6 : 4)}`);
              return `Thread #${t.id} [${t.status ?? "—"}] ${where}\n${comments.join("\n")}`;
            })
            .join("\n\n"),
        );
      }

      case "add_pr_comment": {
        const repo = str(args, "repo");
        const id = num(args, "id");
        const threadId = opt<number>(args, "threadId");
        if (threadId != null) {
          const c = await client.replyToPullRequestThread(repo, id, threadId, str(args, "text"));
          return ok(`Replied in thread #${threadId} (comment ${c.id}).`);
        }
        const filePath = opt<string>(args, "filePath");
        const t = await client.createPullRequestThread(
          repo,
          id,
          str(args, "text"),
          filePath ? { path: filePath, line: opt<number>(args, "line") } : undefined,
        );
        return ok(`Thread #${t.id} created on PR #${id}${filePath ? ` at ${filePath}${opt<number>(args, "line") ? `:${opt<number>(args, "line")}` : ""}` : ""}.`);
      }

      case "update_pr_thread": {
        await client.setPullRequestThreadStatus(
          str(args, "repo"),
          num(args, "id"),
          num(args, "threadId"),
          str(args, "status"),
        );
        return ok(`Thread #${num(args, "threadId")} set to ${str(args, "status")}.`);
      }

      case "vote_pr": {
        const vote = PR_VOTES[str(args, "vote")];
        if (vote == null) return fail(`Unknown vote "${str(args, "vote")}". Use one of: ${Object.keys(PR_VOTES).join(", ")}`);
        await client.setPullRequestVote(str(args, "repo"), num(args, "id"), vote);
        return ok(`Voted ${voteLabel(vote)} on PR #${num(args, "id")}.`);
      }

      case "manage_pr_reviewers": {
        const repo = str(args, "repo");
        const id = num(args, "id");
        const add = opt<string[]>(args, "add") ?? [];
        const remove = opt<string[]>(args, "remove") ?? [];
        if (!add.length && !remove.length) return fail("Pass reviewers to add and/or remove.");
        const required = opt<boolean>(args, "required") ?? false;
        const added = await Promise.all(add.map((r) => client.addPullRequestReviewer(repo, id, r, required)));
        await Promise.all(remove.map((r) => client.removePullRequestReviewer(repo, id, r)));
        return ok(
          [
            added.length ? `Added${required ? " (required)" : ""}: ${added.map((r) => r.displayName).join(", ")}` : "",
            remove.length ? `Removed: ${remove.join(", ")}` : "",
          ]
            .filter(Boolean)
            .join("\n"),
        );
      }

      case "update_pr": {
        const repo = str(args, "repo");
        const id = num(args, "id");
        const action = opt<string>(args, "action");
        const body: Record<string, unknown> = {};
        for (const k of ["title", "description", "isDraft"]) if (args[k] != null) body[k] = args[k];
        const completionOptions = {
          mergeStrategy: opt<string>(args, "mergeStrategy") ?? "squash",
          deleteSourceBranch: opt<boolean>(args, "deleteSourceBranch") ?? true,
          ...(opt<string>(args, "mergeMessage") ? { mergeCommitMessage: opt<string>(args, "mergeMessage") } : {}),
        };
        if (action === "complete") {
          const cur = await client.getPullRequest(repo, id);
          body["status"] = "completed";
          body["lastMergeSourceCommit"] = cur.lastMergeSourceCommit;
          body["completionOptions"] = completionOptions;
        } else if (action === "abandon") {
          body["status"] = "abandoned";
        } else if (action === "reactivate") {
          body["status"] = "active";
        } else if (action === "auto_complete") {
          body["autoCompleteSetBy"] = { id: await client.getMyIdentityId() };
          body["completionOptions"] = completionOptions;
        } else if (action === "cancel_auto_complete") {
          body["autoCompleteSetBy"] = { id: "00000000-0000-0000-0000-000000000000" };
        }
        if (Object.keys(body).length === 0) return fail("Nothing to update. Pass a field to change or an action.");
        const pr = await client.updatePullRequest(repo, id, body);
        return ok(
          [
            `PR #${pr.pullRequestId}: ${pr.title}`,
            `  Status: ${pr.status}${pr.isDraft ? " (draft)" : ""}`,
            `  Merge:  ${pr.mergeStatus}`,
            pr.autoCompleteSetBy ? `  Auto-complete: set by ${pr.autoCompleteSetBy.displayName}` : "",
          ]
            .filter(Boolean)
            .join("\n"),
        );
      }

      case "list_branches": {
        const branches = await client.listBranches(str(args, "repo"));
        if (branches.length === 0) return ok("No branches.");
        return ok(
          branches
            .map((b) =>
              b.isBaseVersion
                ? `${b.name}  (default)`
                : `${b.name}  ↑${b.aheadCount} ↓${b.behindCount}`,
            )
            .join("\n"),
        );
      }

      // ── Repositories ────────────────────────────────────────────────────────
      case "list_repos": {
        const repos = await client.listRepositories();
        if (repos.length === 0) return ok("No repositories found.");
        return ok(
          `${repos.length} repositor${repos.length === 1 ? "y" : "ies"}:\n\n` +
          repos.map((r) => [
            `${r.name}`,
            `  ID:      ${r.id}`,
            `  Branch:  ${r.defaultBranch?.replace("refs/heads/", "") ?? "—"}`,
            `  Size:    ${r.size ? `${Math.round(r.size / 1024)} KB` : "—"}`,
            `  URL:     ${r.remoteUrl}`,
          ].join("\n")).join("\n\n"),
        );
      }

      case "list_commits": {
        const top = pageSize(opt<number>(args, "top") ?? 20);
        const skip = Math.max(0, opt<number>(args, "skip") ?? 0);
        const commits = await client.listCommits(
          str(args, "repo"),
          opt<string>(args, "branch"),
          top,
          skip,
        );
        if (commits.length === 0) return ok(skip > 0 ? "No more commits." : "No commits found.");
        const body = commits.map((c) =>
          `${c.commitId.slice(0, 8)}  ${c.author.date.slice(0, 10)}  ${c.author.name}\n  ${c.comment.split("\n")[0]}`
        ).join("\n");
        return ok(fmtOffsetPage(body, commits.length, skip, top));
      }

      case "get_file": {
        const result = await client.getFileContent(
          str(args, "repo"),
          str(args, "path"),
          opt<string>(args, "branch"),
        );
        if (result.isBinary) return ok(`[Binary file: ${result.path}]`);
        const lines = result.content.split("\n");
        const header = `// ${result.path}${result.commitId ? `  @ ${result.commitId.slice(0, 8)}` : ""}\n\n`;
        // Cap at 500 lines to avoid flooding context
        const body = lines.length > 500
          ? lines.slice(0, 500).join("\n") + `\n\n... (${lines.length - 500} more lines — use a smaller file or specific range)`
          : result.content;
        return ok(header + body);
      }

      case "list_files": {
        const files = await client.listFiles(
          str(args, "repo"),
          opt<string>(args, "path") ?? "/",
          opt<string>(args, "branch"),
        );
        if (files.length === 0) return ok("Empty directory.");
        const dirs = files.filter((f) => f.isFolder);
        const srcs = files.filter((f) => !f.isFolder);
        const fmt = (f: { path: string; isFolder: boolean; size?: number }) => {
          const name = f.path.split("/").pop() ?? f.path;
          return f.isFolder ? `📁 ${name}/` : `   ${name}${f.size ? `  (${Math.round(f.size / 1024) || "<1"} KB)` : ""}`;
        };
        return ok([...dirs, ...srcs].map(fmt).join("\n"));
      }

      // ── Pipelines ────────────────────────────────────────────────────────────
      case "list_pipelines": {
        const pipelines = await client.listPipelines();
        if (pipelines.length === 0) return ok("No pipelines found.");
        return ok(
          `${pipelines.length} pipeline(s):\n\n` +
          pipelines.map((p) =>
            `#${p.id}  ${p.name}${p.folder && p.folder !== "\\" ? `  [${p.folder}]` : ""}`
          ).join("\n"),
        );
      }

      case "run_pipeline": {
        const build = await client.queueBuild(
          num(args, "pipelineId"),
          opt<string>(args, "branch"),
          opt<Record<string, string>>(args, "variables"),
        );
        return ok([
          `Pipeline triggered:`,
          `  Build #${build.id}  ${build.buildNumber}`,
          `  Pipeline: ${build.definition.name}`,
          `  Branch:   ${build.sourceBranch.replace("refs/heads/", "")}`,
          `  Status:   ${build.status}`,
          `  URL:      ${build._links.web.href}`,
        ].join("\n"));
      }

      case "get_build_logs": {
        const buildId = num(args, "buildId");
        const tail = opt<number>(args, "tail");
        const cut = (text: string) => {
          if (tail == null) return text;
          const lines = text.split("\n");
          return lines.length > tail ? `... (${lines.length - tail} lines above)\n${lines.slice(-tail).join("\n")}` : text;
        };
        const logId = opt<number>(args, "logId");
        if (logId != null) return ok(cut(await client.getBuildLogContent(buildId, logId)));

        const records = (await client.getBuildTimeline(buildId)).records ?? [];
        const withLogs = records.filter((r) => r.log?.id != null);
        const step = opt<string>(args, "step")?.toLowerCase();
        const picked: typeof withLogs = step
          ? withLogs.filter((r) => r.name.toLowerCase().includes(step))
          : withLogs.filter((r) => r.result === "failed" && r.type === "Task");
        // A job can fail with no failed task (timeout, agent loss): fall back to any failed record
        if (!step && picked.length === 0) picked.push(...withLogs.filter((r) => r.result === "failed"));
        if (step && picked.length === 0) {
          return fail(`No step matching "${step}". Steps: ${withLogs.map((r) => r.name).join(", ")}`);
        }
        if (picked.length === 0) {
          if (withLogs.length === 0) return ok("No logs available yet.");
          return ok(
            `Build #${buildId}: no failed steps. Step logs (pass logId or step):\n\n` +
              withLogs
                .sort((x, y) => (x.startTime ?? "").localeCompare(y.startTime ?? ""))
                .map((r) => `  log ${r.log!.id}  [${r.type}/${r.result ?? r.state}] ${r.name}`)
                .join("\n"),
          );
        }
        const blocks = await Promise.all(
          picked.map(async (r) => {
            const issues = (r.issues ?? []).map((i) => `  ${i.type}: ${i.message}`).join("\n");
            const text = await client.getBuildLogContent(buildId, r.log!.id);
            return `## ${r.name}  [${r.result ?? r.state}]  log ${r.log!.id}${issues ? `\n${issues}` : ""}\n\n${cut(text)}`;
          }),
        );
        return ok(blocks.join("\n\n"));
      }

      // ── Sprints ──────────────────────────────────────────────────────────────
      case "get_sprint": {
        const teams = await client.listTeams();
        const teamName = opt<string>(args, "team") ?? teams[0]?.name;
        if (!teamName) return ok("No teams found in this project.");
        const [sprint, capacities] = await Promise.all([
          client.getCurrentIteration(teamName),
          client.listTeamCapacities(teamName).catch(() => []),
        ]);
        if (!sprint) return ok(`No active sprint for team "${teamName}".`);
        const lines = [
          `Sprint: ${sprint.name}`,
          `Team:   ${teamName}`,
          `Start:  ${sprint.attributes.startDate?.slice(0, 10) ?? "—"}`,
          `End:    ${sprint.attributes.finishDate?.slice(0, 10) ?? "—"}`,
        ];
        if (capacities.length) {
          lines.push("", "Capacity:");
          for (const c of capacities) {
            const total = c.activities.reduce((s, a) => s + a.capacityPerDay, 0);
            lines.push(`  ${c.teamMember.displayName}  ${total}h/day`);
          }
        }
        return ok(lines.join("\n"));
      }

      case "list_sprints": {
        const teams = await client.listTeams();
        const teamName = opt<string>(args, "team") ?? teams[0]?.name;
        if (!teamName) return ok("No teams found in this project.");
        const iterations = await client.listIterations(teamName);
        if (iterations.length === 0) return ok("No sprints found.");
        return ok(
          `Sprints for "${teamName}" (${iterations.length}):\n\n` +
          iterations.map((it) => {
            const start = it.attributes.startDate?.slice(0, 10) ?? "—";
            const end = it.attributes.finishDate?.slice(0, 10) ?? "—";
            return `${it.name}  (${start} → ${end})`;
          }).join("\n"),
        );
      }

      // ── Backlog ────────────────────────────────────────────────────────────
      case "get_backlog": {
        const teams = await client.listTeams();
        const teamName = opt<string>(args, "team") ?? teams[0]?.name;
        if (!teamName) return ok("No teams found in this project.");
        const top = pageSize(opt<number>(args, "top") ?? 20);
        const skip = Math.max(0, opt<number>(args, "skip") ?? 0);
        const all = await client.getBacklog(teamName, opt<string>(args, "level"));
        if (all.length === 0) return ok("Backlog is empty.");
        const page = all.slice(skip, skip + top);
        const body = page
          .map((b, i) => {
            const rank = skip + i + 1;
            const meta = [
              b.assignedTo !== "Unassigned" ? `@${b.assignedTo}` : null,
              b.storyPoints != null ? `${b.storyPoints}pts` : null,
              b.priority != null ? `P${b.priority}` : null,
            ].filter(Boolean).join(" · ");
            return `${rank}. #${b.id} [${b.type}/${b.state}] ${b.title}${meta ? `  — ${meta}` : ""}`;
          })
          .join("\n");
        const more = skip + page.length < all.length
          ? `\n\n— ${all.length - skip - page.length} more. Call again with skip: ${skip + page.length}.`
          : "";
        return ok(`Backlog for "${teamName}" — ${all.length} item(s):\n\n${body}${more}`);
      }

      // ── Work item history ──────────────────────────────────────────────────
      case "get_work_item_history": {
        const top = opt<number>(args, "top") ?? 15;
        const updates = await client.getWorkItemHistory(num(args, "id"));
        // Most recent first, keep only revisions that actually changed a field.
        const meaningful = updates
          .filter((u) => Object.keys(u.fields).length > 0)
          .reverse()
          .slice(0, top);
        if (meaningful.length === 0) return ok("No change history.");
        const SKIP = new Set([
          "System.Rev", "System.AuthorizedDate", "System.RevisedDate",
          "System.ChangedDate", "System.Watermark", "System.AuthorizedAs",
          "System.PersonId", "System.ChangedBy",
        ]);
        const short = (k: string) => k.split(".").pop() ?? k;
        const val = (v: unknown) => {
          if (v == null || v === "") return "—";
          return typeof v === "object" && v && "displayName" in v
            ? String((v as { displayName: string }).displayName)
            : stripHtml(String(v));
        };
        const blocks = meaningful.map((u) => {
          const who = u.revisedBy?.displayName ?? "—";
          // revisedDate is when this revision was superseded (9999 for the latest); ChangedDate is when it happened
          const changed = u.fields["System.ChangedDate"]?.newValue;
          const when = String(changed ?? u.revisedDate ?? "—").slice(0, 16).replace("T", " ");
          const changes = Object.entries(u.fields)
            .filter(([k]) => !SKIP.has(k))
            .map(([k, d]) => `  ${short(k)}: ${val(d.oldValue)} → ${val(d.newValue)}`);
          return changes.length
            ? `[${when}] ${who}\n${changes.join("\n")}`
            : null;
        }).filter(Boolean);
        return ok(blocks.length ? blocks.join("\n\n") : "No field changes in recent history.");
      }

      // ── Teams & paths ──────────────────────────────────────────────────────
      case "list_team_members": {
        const teams = await client.listTeams();
        const teamName = opt<string>(args, "team") ?? teams[0]?.name;
        if (!teamName) return ok("No teams found in this project.");
        const members = await client.getTeamMembers(teamName);
        if (members.length === 0) return ok(`No members in team "${teamName}".`);
        return ok(
          `Team "${teamName}" — ${members.length} member(s):\n\n` +
          members.map((m) =>
            `${m.identity.displayName}${m.isTeamAdmin ? " (admin)" : ""}  <${m.identity.uniqueName}>`
          ).join("\n"),
        );
      }

      case "list_paths": {
        const kind = str(args, "kind");
        if (kind === "iteration") {
          const paths = await client.listIterationPaths();
          if (paths.length === 0) return ok("No iteration paths found.");
          return ok(
            `Iteration paths (${paths.length}):\n\n` +
            paths.map((p) => {
              const dates = p.startDate || p.finishDate
                ? `  (${p.startDate?.slice(0, 10) ?? "?"} → ${p.finishDate?.slice(0, 10) ?? "?"})`
                : "";
              return `${p.path}${dates}`;
            }).join("\n"),
          );
        }
        const paths = await client.listAreaPaths();
        if (paths.length === 0) return ok("No area paths found.");
        return ok(`Area paths (${paths.length}):\n\n` + paths.map((p) => p.path).join("\n"));
      }

      case "get_board": {
        const teams = await client.listTeams();
        const teamName = opt<string>(args, "team") ?? teams[0]?.name;
        if (!teamName) return ok("No teams found in this project.");
        const boards = await client.listBoards(teamName);
        if (boards.length === 0) return ok(`No boards for team "${teamName}".`);
        const wanted = opt<string>(args, "board");
        const board = wanted
          ? boards.find((b) => b.name.toLowerCase() === wanted.toLowerCase() || b.id === wanted) ?? boards[0]!
          : boards[0]!;
        const detail = await client.getBoard(teamName, board.id);
        return ok(
          `Board "${detail.name}" (team ${teamName}):\n\n` +
          detail.columns.map((c) => {
            const states = Object.entries(c.stateMappings).map(([wit, st]) => `${wit}→${st}`).join(", ");
            const limit = c.itemLimit > 0 ? `  [WIP ${c.itemLimit}]` : "";
            return `│ ${c.name}${limit}${states ? `\n│   ${states}` : ""}`;
          }).join("\n"),
        );
      }

      // ── Build timeline & cancel ──────────────────────────────────────────────
      case "get_build_timeline": {
        const timeline = await client.getBuildTimeline(num(args, "buildId"));
        const records = timeline.records ?? [];
        if (records.length === 0) return ok("No timeline available (build may not have started).");
        const icon = (r: { state: string; result: string | null }) =>
          r.result === "succeeded" ? "✅"
          : r.result === "failed" ? "❌"
          : r.result === "canceled" ? "⏹"
          : r.state === "inProgress" ? "🔄"
          : "○";
        // Sort by type depth (Stage → Phase → Job → Task) then start time.
        const order: Record<string, number> = { Stage: 0, Phase: 1, Job: 2, Task: 3 };
        const sorted = [...records].sort((a, b) =>
          (order[a.type] ?? 9) - (order[b.type] ?? 9) ||
          (a.startTime ?? "").localeCompare(b.startTime ?? ""),
        );
        const lines = sorted.map((r) => {
          const dur = r.startTime && r.finishTime
            ? `${Math.round((new Date(r.finishTime).getTime() - new Date(r.startTime).getTime()) / 1000)}s`
            : "";
          const issues = [
            r.errorCount ? `${r.errorCount} err` : null,
            r.warningCount ? `${r.warningCount} warn` : null,
          ].filter(Boolean).join(", ");
          const indent = "  ".repeat(order[r.type] ?? 0);
          return `${icon(r)} ${indent}${r.name}${dur ? `  (${dur})` : ""}${issues ? `  [${issues}]` : ""}`;
        });
        const failed = sorted.filter((r) => r.result === "failed");
        const header = failed.length
          ? `Build #${num(args, "buildId")} — ${failed.length} failed step(s): ${failed.map((f) => f.name).join(", ")}\n\n`
          : `Build #${num(args, "buildId")} timeline:\n\n`;
        return ok(header + lines.join("\n"));
      }

      case "cancel_build": {
        await client.cancelBuild(num(args, "buildId"));
        return ok(`Build #${num(args, "buildId")} cancellation requested.`);
      }

      // ── Releases ─────────────────────────────────────────────────────────────
      case "list_releases": {
        const releases = await client.listReleases({
          definitionId: opt<number>(args, "definitionId"),
          status: opt<"draft" | "active" | "abandoned">(args, "status"),
          top: opt<number>(args, "top") ?? 10,
        });
        if (releases.length === 0) return ok("No releases found.");
        return ok(
          releases.map((r) => {
            const envs = r.environments.map((e) => {
              const i = e.status === "succeeded" ? "✅"
                : e.status === "failed" || e.status === "rejected" ? "❌"
                : e.status === "inProgress" ? "🔄"
                : e.status === "notStarted" ? "○"
                : "·";
              return `${i} ${e.name}`;
            }).join("  ");
            return [
              `#${r.id} ${r.name} [${r.status}]`,
              `  Def:  ${r.releaseDefinition.name}`,
              `  By:   ${r.createdBy.displayName} · ${r.createdOn.slice(0, 10)}`,
              `  Envs: ${envs || "—"}`,
              `  URL:  ${r._links.web.href}`,
            ].join("\n");
          }).join("\n\n"),
        );
      }

      case "create_release": {
        const r = await client.createRelease(
          num(args, "definitionId"),
          opt<string>(args, "description"),
        );
        return ok([
          `Release created:`,
          `  #${r.id}  ${r.name}`,
          `  Definition: ${r.releaseDefinition.name}`,
          `  Status:     ${r.status}`,
          `  URL:        ${r._links.web.href}`,
        ].join("\n"));
      }

      default:
        return { content: [{ type: "text", text: `Unknown tool: ${name}` }], isError: true };
    }
  } catch (e) {
    return fail(e);
  }
}
