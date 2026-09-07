export function getTableFlowFragments(
  root: ParentNode | null,
  table: HTMLTableElement
): HTMLTableElement[] {
  const flowId = table.getAttribute("data-hwe-table-flow-id");
  if (!root || !flowId) return [table];

  const fragments = Array.from(root.querySelectorAll<HTMLTableElement>("table")).filter(
    (candidate) => candidate.getAttribute("data-hwe-table-flow-id") === flowId
  );
  return fragments.length > 0 ? fragments : [table];
}

export function getTableFlowRows(
  root: ParentNode | null,
  table: HTMLTableElement
): HTMLTableRowElement[] {
  return getTableFlowFragments(root, table).flatMap((fragment) => Array.from(fragment.rows));
}
