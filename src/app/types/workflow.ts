import type { RowData } from '../calculatorDomain'

export type WorkflowReviewer = {
  id: string
  username: string
  role: string
}

export type PerformanceWorkflowItem = {
  id: string
  period: string
  status: string
  subsidiaryLabel: string
  storeLabel: string
  rowCountCalculated: number
  rowCountUploaded: number
  assignedToUserName: string
  submittedByName: string
  pushedAt: string
  confirmedAt: string
  archivedAt: string
  createdAt: string
  uploadedRows: RowData[]
}
