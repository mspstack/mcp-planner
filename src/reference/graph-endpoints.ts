/**
 * Curated Microsoft Graph endpoint catalog for the advanced toolset
 * (graph_find_endpoint / graph_get).
 *
 * Hand-maintained: Graph's full OpenAPI is enormous and almost all of it is
 * unreachable with this server's app permissions (Tasks.ReadWrite.All,
 * GroupMember.Read.All, User.Read.All), so the catalog lists only the honest
 * surface — the /planner, /groups and /users paths graph_get allows. Keep
 * entries alphabetized by path (after the synthetic grammar entry).
 */

export interface EndpointDoc {
  module: string;
  path: string;
  methods: string;
  summary: string;
  /** Curated tool(s) that already wrap this endpoint — prefer them. */
  coveredBy?: string;
  /** Most useful query params. */
  keyParams?: string;
  /** Common properties in the response. */
  commonFields?: string;
  notes?: string;
}

export const GRAPH_ENDPOINTS: EndpointDoc[] = [
  {
    module: "meta",
    path: "(query grammar)",
    methods: "—",
    summary:
      "How Microsoft Graph OData queries work: $select (comma-separated properties — always pass it to keep " +
      "responses small), $filter (e.g. startswith(displayName,'NDR')), $orderby, $top, $expand, $search " +
      "(directory objects only, value like \"displayName:pod\", needs ConsistencyLevel: eventual — graph_get adds " +
      "the header automatically). CAVEAT: planner collections (/planner/*) ignore $filter and $top server-side — " +
      "filter client-side and use $select to slim payloads. Paging: responses carry @odata.nextLink; pass the " +
      "whole nextLink back as the path.",
    notes: "Query syntax reference — not a callable endpoint.",
  },
  {
    module: "groups",
    path: "/groups",
    methods: "GET",
    summary: "Microsoft 365 groups (each group can own Planner plans).",
    coveredBy: "planner_search_groups",
    keyParams: "$filter (e.g. startswith(displayName,'x')), $search (\"displayName:x\"), $select=id,displayName,mail",
    commonFields: "id, displayName, mail, description, groupTypes, visibility",
    notes: "$search requires ConsistencyLevel: eventual (added automatically).",
  },
  {
    module: "groups",
    path: "/groups/delta",
    methods: "GET",
    summary: "Track incremental changes to groups (delta query).",
    keyParams: "$select; follow @odata.nextLink / @odata.deltaLink",
  },
  {
    module: "groups",
    path: "/groups/{group_id}",
    methods: "GET",
    summary: "One group's profile.",
    keyParams: "$select",
    commonFields: "id, displayName, mail, description, createdDateTime",
  },
  {
    module: "groups",
    path: "/groups/{group_id}/members",
    methods: "GET",
    summary: "Members of a group (directoryObjects — usually users).",
    keyParams: "$select=id,displayName,mail, $top, $count=true",
    commonFields: "id, displayName, mail, userPrincipalName",
  },
  {
    module: "groups",
    path: "/groups/{group_id}/owners",
    methods: "GET",
    summary: "Owners of a group.",
    keyParams: "$select=id,displayName,mail",
  },
  {
    module: "groups",
    path: "/groups/{group_id}/planner/plans",
    methods: "GET",
    summary: "Planner plans owned by a group.",
    coveredBy: "planner_list_plans",
    commonFields: "id, title, owner, createdDateTime, createdBy",
    notes: "Planner collection — no server-side $filter/$top.",
  },
  {
    module: "planner",
    path: "/planner/buckets/{bucket_id}",
    methods: "GET",
    summary: "One Planner bucket (column).",
    commonFields: "id, name, planId, orderHint, @odata.etag",
  },
  {
    module: "planner",
    path: "/planner/buckets/{bucket_id}/tasks",
    methods: "GET",
    summary: "Tasks in one bucket.",
    coveredBy: "planner_list_tasks",
    commonFields: "id, title, percentComplete, priority, dueDateTime, assignments, @odata.etag",
    notes: "Planner collection — no server-side $filter/$top.",
  },
  {
    module: "planner",
    path: "/planner/plans/{plan_id}",
    methods: "GET",
    summary: "One Planner plan.",
    coveredBy: "planner_get_plan",
    commonFields: "id, title, owner (group id), createdDateTime, createdBy, @odata.etag",
  },
  {
    module: "planner",
    path: "/planner/plans/{plan_id}/buckets",
    methods: "GET",
    summary: "Buckets (columns) of a plan.",
    coveredBy: "planner_get_plan",
    commonFields: "id, name, planId, orderHint",
    notes: "Planner collection — no server-side $filter/$top.",
  },
  {
    module: "planner",
    path: "/planner/plans/{plan_id}/details",
    methods: "GET",
    summary: "Plan details: category descriptions (label names) and shared-with list.",
    commonFields: "categoryDescriptions (category1–25), sharedWith, @odata.etag",
  },
  {
    module: "planner",
    path: "/planner/plans/{plan_id}/tasks",
    methods: "GET",
    summary: "All tasks of a plan.",
    coveredBy: "planner_list_tasks",
    commonFields: "id, title, bucketId, percentComplete, priority, dueDateTime, assignments, @odata.etag",
    notes: "Planner collection — no server-side $filter/$top.",
  },
  {
    module: "planner",
    path: "/planner/tasks/{task_id}",
    methods: "GET",
    summary: "One Planner task (assignments, dates, progress, applied categories).",
    coveredBy: "planner_get_task",
    commonFields:
      "id, title, planId, bucketId, percentComplete (0/50/100), priority (1=urgent, 3=important, 5=medium, 9=low), " +
      "startDateTime, dueDateTime, completedDateTime, assignments, appliedCategories, @odata.etag",
    notes: "Writes need If-Match with the current @odata.etag — use the curated planner_update_task.",
  },
  {
    module: "planner",
    path: "/planner/tasks/{task_id}/assignedToTaskBoardFormat",
    methods: "GET",
    summary: "Task ordering hints on the 'Assigned to' board view.",
    commonFields: "orderHintsByAssignee, unassignedOrderHint, @odata.etag",
  },
  {
    module: "planner",
    path: "/planner/tasks/{task_id}/bucketTaskBoardFormat",
    methods: "GET",
    summary: "Task ordering hint on the bucket board view.",
    commonFields: "orderHint, @odata.etag",
  },
  {
    module: "planner",
    path: "/planner/tasks/{task_id}/details",
    methods: "GET",
    summary: "Task details: description, checklist, references (attachments/links), preview type.",
    coveredBy: "planner_get_task",
    commonFields: "description, checklist (item id → {title, isChecked}), references, previewType, @odata.etag",
    notes: "Separate resource with its OWN etag, distinct from the task's.",
  },
  {
    module: "planner",
    path: "/planner/tasks/{task_id}/progressTaskBoardFormat",
    methods: "GET",
    summary: "Task ordering hint on the progress board view.",
    commonFields: "orderHint, @odata.etag",
  },
  {
    module: "users",
    path: "/users",
    methods: "GET",
    summary: "Directory users.",
    coveredBy: "planner_find_user",
    keyParams: "$filter (e.g. startswith(displayName,'x') or mail eq 'x@y.com'), $search, $select=id,displayName,mail,userPrincipalName",
    commonFields: "id, displayName, mail, userPrincipalName, jobTitle, department, accountEnabled",
    notes: "App-only auth cannot use /me — address users by id or userPrincipalName.",
  },
  {
    module: "users",
    path: "/users/delta",
    methods: "GET",
    summary: "Track incremental changes to users (delta query).",
    keyParams: "$select; follow @odata.nextLink / @odata.deltaLink",
  },
  {
    module: "users",
    path: "/users/{user_id}",
    methods: "GET",
    summary: "One user's profile (id or userPrincipalName).",
    keyParams: "$select",
    commonFields: "id, displayName, mail, userPrincipalName, jobTitle, department, officeLocation",
  },
  {
    module: "users",
    path: "/users/{user_id}/memberOf",
    methods: "GET",
    summary: "Groups (and directory roles) the user is a member of.",
    keyParams: "$select=id,displayName, $count=true",
  },
  {
    module: "users",
    path: "/users/{user_id}/planner/tasks",
    methods: "GET",
    summary: "All Planner tasks assigned to a user, across every plan.",
    coveredBy: "planner_list_user_tasks",
    commonFields: "id, title, planId, bucketId, percentComplete, dueDateTime, @odata.etag",
    notes: "Planner collection — no server-side $filter/$top.",
  },
];
