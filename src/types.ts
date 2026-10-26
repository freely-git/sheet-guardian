export type PermissionRole = 'reader' | 'commenter' | 'writer'

export interface PermissionDetail {
  inherited?: boolean
  inheritedFrom?: string
  permissionType?: string
  role?: string
}

export interface DrivePermission {
  id: string
  type: 'user' | 'group' | 'domain' | 'anyone'
  role: string
  emailAddress?: string
  displayName?: string
  domain?: string
  allowFileDiscovery?: boolean
  deleted?: boolean
  pendingOwner?: boolean
  permissionDetails?: PermissionDetail[]
}

export interface SheetFile {
  id: string
  name: string
  mimeType: string
  webViewLink?: string
  modifiedTime?: string
  createdTime?: string
  ownedByMe?: boolean
  driveId?: string
  parents?: string[]
  inheritedPermissionsDisabled?: boolean
  trashed?: boolean
  size?: string
  owners?: Array<{ displayName?: string; emailAddress?: string }>
  capabilities?: { canShare?: boolean; canTrash?: boolean; canDelete?: boolean; canDisableInheritedPermissions?: boolean }
  permissions?: DrivePermission[]
}

export interface DriveFolder {
  id: string
  name: string
  mimeType: string
  webViewLink?: string
  driveId?: string
  parents?: string[]
  inheritedPermissionsDisabled?: boolean
  capabilities?: { canDisableInheritedPermissions?: boolean }
}

export interface UserProfile {
  email: string
  id: string
}

export type LogStatus = 'success' | 'failed' | 'skipped'

export interface OperationLog {
  id: string
  time: string
  action: '授权' | 'Sheet 保护权限' | '转移所有权' | '开启共享' | '收紧共享' | '移至回收站'
  status: LogStatus
  fileId: string
  fileName: string
  target?: string
  detail: string
}

export interface ProtectedSheetEditorResult {
  matchedSheets: number
  sheetsWithProtection: number
  eligibleRanges: number
  updatedRanges: number
  alreadyPresentRanges: number
  skippedWarningOnly: number
  skippedUneditable: number
}

export interface GrantResult {
  status: 'created' | 'updated' | 'skipped'
  permission?: DrivePermission
  message: string
}

export interface RevokeResult {
  removed: DrivePermission[]
  skippedInherited: DrivePermission[]
}
