import type {
  DriveFolder,
  DrivePermission,
  GrantResult,
  PermissionRole,
  ProtectedSheetEditorResult,
  RevokeResult,
  SheetFile,
  UserProfile,
} from '../types'

const DRIVE_API = 'https://www.googleapis.com/drive/v3'
const SHEETS_API = 'https://sheets.googleapis.com/v4'
const FOLDER_MIME = 'application/vnd.google-apps.folder'
let accessToken = ''

type ChromeIdentityResult = { token?: string; grantedScopes?: string[] }

export function hasChromeIdentity(): boolean {
  return typeof chrome !== 'undefined' && Boolean(chrome.identity?.getAuthToken)
}

function lastErrorMessage(): string | undefined {
  return chrome.runtime?.lastError?.message
}

function requestAuthToken(interactive: boolean): Promise<string> {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (result?: string | ChromeIdentityResult) => {
      const error = lastErrorMessage()
      if (error) {
        reject(new Error(error))
        return
      }
      const token = typeof result === 'string' ? result : result?.token
      if (!token) {
        reject(new Error('Google 未返回访问令牌'))
        return
      }
      resolve(token)
    })
  })
}

function requestProfile(): Promise<UserProfile> {
  return new Promise((resolve) => {
    chrome.identity.getProfileUserInfo({ accountStatus: chrome.identity.AccountStatus.ANY }, (info) => {
      resolve({ email: info.email || '', id: info.id || '' })
    })
  })
}

function removeToken(token: string): Promise<void> {
  return new Promise((resolve) => {
    chrome.identity.removeCachedAuthToken({ token }, () => resolve())
  })
}

export async function connectGoogle(interactive = true): Promise<UserProfile> {
  if (!hasChromeIdentity()) throw new Error('当前页面不是 Chrome 扩展环境')
  accessToken = await requestAuthToken(interactive)
  return requestProfile()
}

export async function disconnectGoogle(): Promise<void> {
  if (accessToken) await removeToken(accessToken)
  accessToken = ''
  if (hasChromeIdentity()) await new Promise<void>((resolve) => chrome.identity.clearAllCachedAuthTokens(() => resolve()))
}

function errorMessage(payload: any, status: number): string {
  return payload?.error?.message || payload?.message || `Google API 请求失败 (${status})`
}

async function driveRequest<T>(path: string, init: RequestInit = {}, retry = true): Promise<T> {
  if (!accessToken) accessToken = await requestAuthToken(false)

  const response = await fetch(`${DRIVE_API}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  })

  if (response.status === 401 && retry) {
    await removeToken(accessToken)
    accessToken = await requestAuthToken(true)
    return driveRequest<T>(path, init, false)
  }

  if (!response.ok) {
    const payload = await response.json().catch(() => ({}))
    const error = new Error(errorMessage(payload, response.status)) as Error & { status?: number }
    error.status = response.status
    throw error
  }

  if (response.status === 204) return undefined as T
  return response.json() as Promise<T>
}

async function sheetsRequest<T>(path: string, init: RequestInit = {}, retry = true): Promise<T> {
  if (!accessToken) accessToken = await requestAuthToken(false)

  const response = await fetch(`${SHEETS_API}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  })

  if (response.status === 401 && retry) {
    await removeToken(accessToken)
    accessToken = await requestAuthToken(true)
    return sheetsRequest<T>(path, init, false)
  }

  if (!response.ok) {
    const payload = await response.json().catch(() => ({}))
    let message = errorMessage(payload, response.status)
    if (/sheets api|api has not been used|access not configured/i.test(message)) {
      message = `${message}；请在 Google Cloud 项目中启用 Google Sheets API`
    }
    const error = new Error(message) as Error & { status?: number }
    error.status = response.status
    throw error
  }

  if (response.status === 204) return undefined as T
  return response.json() as Promise<T>
}

type ProtectedRange = {
  protectedRangeId?: number
  warningOnly?: boolean
  requestingUserCanEdit?: boolean
  editors?: {
    users?: string[]
    groups?: string[]
    domainUsersCanEdit?: boolean
  }
}

type SpreadsheetSheet = {
  properties?: { sheetId?: number; title?: string }
  protectedRanges?: ProtectedRange[]
}

export async function addProtectedSheetEditors(
  spreadsheetId: string,
  keywords: string[],
  emails: string[],
): Promise<ProtectedSheetEditorResult> {
  const normalizedKeywords = [...new Set(keywords.map((keyword) => keyword.trim().toLowerCase()).filter(Boolean))]
  const normalizedEmails = [...new Set(emails.map((email) => email.trim().toLowerCase()).filter(Boolean))]
  const result: ProtectedSheetEditorResult = {
    matchedSheets: 0,
    sheetsWithProtection: 0,
    eligibleRanges: 0,
    updatedRanges: 0,
    alreadyPresentRanges: 0,
    skippedWarningOnly: 0,
    skippedUneditable: 0,
  }
  if (!normalizedKeywords.length || !normalizedEmails.length) return result

  const params = new URLSearchParams({
    fields:
      'sheets(properties(sheetId,title),protectedRanges(protectedRangeId,warningOnly,requestingUserCanEdit,editors(users,groups,domainUsersCanEdit)))',
  })
  const spreadsheet = await sheetsRequest<{ sheets?: SpreadsheetSheet[] }>(
    `/spreadsheets/${encodeURIComponent(spreadsheetId)}?${params}`,
  )
  const requests: Array<{
    updateProtectedRange: {
      protectedRange: { protectedRangeId: number; editors: { users: string[] } }
      fields: 'editors.users'
    }
  }> = []

  for (const sheet of spreadsheet.sheets || []) {
    const title = sheet.properties?.title?.trim() || ''
    if (!normalizedKeywords.some((keyword) => title.toLowerCase().includes(keyword))) continue
    result.matchedSheets++

    const protectedRanges = sheet.protectedRanges || []
    if (protectedRanges.length) result.sheetsWithProtection++

    for (const protectedRange of protectedRanges) {
      if (protectedRange.warningOnly) {
        result.skippedWarningOnly++
        continue
      }
      if (protectedRange.requestingUserCanEdit !== true || protectedRange.protectedRangeId === undefined) {
        result.skippedUneditable++
        continue
      }

      result.eligibleRanges++
      const existingUsers = protectedRange.editors?.users || []
      const existingUserKeys = new Set(existingUsers.map((email) => email.toLowerCase()))
      const usersToAdd = normalizedEmails.filter((email) => !existingUserKeys.has(email))
      if (!usersToAdd.length) {
        result.alreadyPresentRanges++
        continue
      }

      requests.push({
        updateProtectedRange: {
          protectedRange: {
            protectedRangeId: protectedRange.protectedRangeId,
            editors: { users: [...existingUsers, ...usersToAdd] },
          },
          fields: 'editors.users',
        },
      })
    }
  }

  if (requests.length) {
    await sheetsRequest(`/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({ requests }),
    })
    result.updatedRanges = requests.length
  }

  return result
}

const FILE_FIELDS = [
  'id',
  'name',
  'mimeType',
  'webViewLink',
  'modifiedTime',
  'createdTime',
  'size',
  'ownedByMe',
  'driveId',
  'parents',
  'inheritedPermissionsDisabled',
  'owners(displayName,emailAddress)',
  'capabilities(canShare,canTrash,canDelete,canDisableInheritedPermissions)',
  'permissions(id,type,role,emailAddress,displayName,domain,allowFileDiscovery,deleted,pendingOwner,permissionDetails(inherited,inheritedFrom,permissionType,role))',
].join(',')

export async function listDriveFiles(onPage?: (count: number) => void): Promise<SheetFile[]> {
  const files: SheetFile[] = []
  let pageToken = ''

  do {
    const params = new URLSearchParams({
      q: 'trashed=false',
      spaces: 'drive',
      corpora: 'allDrives',
      includeItemsFromAllDrives: 'true',
      supportsAllDrives: 'true',
      orderBy: 'modifiedTime desc',
      pageSize: '1000',
      fields: `nextPageToken,files(${FILE_FIELDS})`,
    })
    if (pageToken) params.set('pageToken', pageToken)

    const page = await driveRequest<{ files?: SheetFile[]; nextPageToken?: string }>(`/files?${params}`)
    files.push(...(page.files || []))
    pageToken = page.nextPageToken || ''
    onPage?.(files.length)
  } while (pageToken)

  return files
}

export async function getDriveFile(fileId: string): Promise<SheetFile> {
  const params = new URLSearchParams({ fields: FILE_FIELDS, supportsAllDrives: 'true' })
  return driveRequest<SheetFile>(`/files/${encodeURIComponent(fileId)}?${params}`)
}

export async function trashSpreadsheet(fileId: string): Promise<SheetFile> {
  const params = new URLSearchParams({
    fields: 'id,name,mimeType,trashed',
    supportsAllDrives: 'true',
  })
  return driveRequest<SheetFile>(`/files/${encodeURIComponent(fileId)}?${params}`, {
    method: 'PATCH',
    body: JSON.stringify({ trashed: true }),
  })
}

const FOLDER_FIELDS = [
  'id',
  'name',
  'mimeType',
  'webViewLink',
  'driveId',
  'parents',
  'inheritedPermissionsDisabled',
  'capabilities(canDisableInheritedPermissions)',
].join(',')

export async function enableFolderLimitedAccess(folderId: string): Promise<DriveFolder> {
  const getParams = new URLSearchParams({ fields: FOLDER_FIELDS, supportsAllDrives: 'true' })
  const folder = await driveRequest<DriveFolder>(`/files/${encodeURIComponent(folderId)}?${getParams}`)
  if (folder.mimeType !== FOLDER_MIME) throw new Error('权限来源不是可设置有限访问的文件夹')
  if (folder.inheritedPermissionsDisabled) return folder
  if (folder.capabilities?.canDisableInheritedPermissions !== true) {
    throw new Error(`当前账号无权为文件夹“${folder.name}”启用有限访问`)
  }

  const updateParams = new URLSearchParams({ fields: FOLDER_FIELDS, supportsAllDrives: 'true' })
  return driveRequest<DriveFolder>(`/files/${encodeURIComponent(folderId)}?${updateParams}`, {
    method: 'PATCH',
    body: JSON.stringify({ inheritedPermissionsDisabled: true }),
  })
}

export async function listPermissions(fileId: string): Promise<DrivePermission[]> {
  const params = new URLSearchParams({
    supportsAllDrives: 'true',
    fields:
      'permissions(id,type,role,emailAddress,displayName,domain,allowFileDiscovery,deleted,pendingOwner,permissionDetails(inherited,inheritedFrom,permissionType,role))',
  })
  const result = await driveRequest<{ permissions?: DrivePermission[] }>(
    `/files/${encodeURIComponent(fileId)}/permissions?${params}`,
  )
  return result.permissions || []
}

export function isInheritedPermission(permission: DrivePermission): boolean {
  return Boolean(permission.permissionDetails?.some((detail) => detail.inherited || detail.inheritedFrom))
}

function isInheritedPermissionError(error: unknown): boolean {
  if (!(error instanceof Error)) return false
  const message = error.message.toLowerCase()
  return (
    message.includes('inherited permission') ||
    message.includes('cannotmodifyinheritedteamdrivepermission') ||
    (message.includes('cannot delete the permission') && message.includes('limited access'))
  )
}

export async function tightenSharing(
  fileId: string,
  types: Array<'anyone' | 'domain'>,
  _knownPermissions?: DrivePermission[],
): Promise<RevokeResult> {
  // Always refresh immediately before deletion. File-list results can contain a
  // reduced or stale permissionDetails payload, which makes inherited ACLs look direct.
  const permissions = await listPermissions(fileId)
  const candidates = permissions.filter((permission) => types.includes(permission.type as 'anyone' | 'domain'))
  const inherited = candidates.filter(isInheritedPermission)
  const removable = candidates.filter((permission) => !isInheritedPermission(permission))
  const removed: DrivePermission[] = []

  for (const permission of removable) {
    try {
      const params = new URLSearchParams({ supportsAllDrives: 'true' })
      await driveRequest<void>(
        `/files/${encodeURIComponent(fileId)}/permissions/${encodeURIComponent(permission.id)}?${params}`,
        { method: 'DELETE' },
      )
      removed.push(permission)
    } catch (error) {
      // Google can still identify an ACL as inherited only at deletion time.
      // Limited access is a folder setting, so spreadsheets must skip it and
      // direct the user to the originating parent folder/shared drive instead.
      if (isInheritedPermissionError(error)) {
        inherited.push(permission)
        continue
      }
      throw error
    }
  }

  return { removed, skippedInherited: inherited }
}

export async function grantPermission(
  fileId: string,
  email: string,
  role: PermissionRole,
  notify: boolean,
  emailMessage?: string,
  knownPermissions?: DrivePermission[],
): Promise<GrantResult> {
  const permissions = knownPermissions?.length ? knownPermissions : await listPermissions(fileId)
  const existing = permissions.find(
    (permission) =>
      ['user', 'group'].includes(permission.type) && permission.emailAddress?.toLowerCase() === email.toLowerCase(),
  )

  if (existing?.role === 'owner') {
    return { status: 'skipped', permission: existing, message: '该账号是所有者，已保留' }
  }
  if (existing?.role === role) {
    return { status: 'skipped', permission: existing, message: '权限已存在' }
  }
  if (existing) {
    const params = new URLSearchParams({ supportsAllDrives: 'true', fields: 'id,type,role,emailAddress,displayName' })
    const permission = await driveRequest<DrivePermission>(
      `/files/${encodeURIComponent(fileId)}/permissions/${encodeURIComponent(existing.id)}?${params}`,
      { method: 'PATCH', body: JSON.stringify({ role }) },
    )
    return { status: 'updated', permission, message: `权限已更新为 ${role}` }
  }

  const params = new URLSearchParams({
    supportsAllDrives: 'true',
    sendNotificationEmail: String(notify),
    fields: 'id,type,role,emailAddress,displayName',
  })
  if (notify && emailMessage?.trim()) params.set('emailMessage', emailMessage.trim())

  const permission = await driveRequest<DrivePermission>(`/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
    method: 'POST',
    body: JSON.stringify({ type: 'user', role, emailAddress: email }),
  })
  return { status: 'created', permission, message: '权限已添加' }
}

export async function enableAnyoneLinkSharing(
  fileId: string,
  role: PermissionRole,
  knownPermissions?: DrivePermission[],
): Promise<GrantResult> {
  const permissions = knownPermissions?.length ? knownPermissions : await listPermissions(fileId)
  const existing = permissions.find((permission) => permission.type === 'anyone')

  if (existing) {
    if (existing.role === role) {
      return { status: 'skipped', permission: existing, message: `链接共享已经是 ${role}` }
    }
    if (isInheritedPermission(existing)) {
      return {
        status: 'skipped',
        permission: existing,
        message: `链接共享继承自上级位置，已保留现有 ${existing.role} 权限`,
      }
    }

    const params = new URLSearchParams({ supportsAllDrives: 'true', fields: 'id,type,role,allowFileDiscovery' })
    const permission = await driveRequest<DrivePermission>(
      `/files/${encodeURIComponent(fileId)}/permissions/${encodeURIComponent(existing.id)}?${params}`,
      { method: 'PATCH', body: JSON.stringify({ role }) },
    )
    return { status: 'updated', permission, message: `链接共享已更新为 ${role}` }
  }

  const params = new URLSearchParams({
    supportsAllDrives: 'true',
    sendNotificationEmail: 'false',
    fields: 'id,type,role,allowFileDiscovery',
  })
  const permission = await driveRequest<DrivePermission>(`/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
    method: 'POST',
    body: JSON.stringify({ type: 'anyone', role, allowFileDiscovery: false }),
  })
  return { status: 'created', permission, message: `已开启知道链接的任何人（${role}）` }
}

function isConsumerGmail(email: string): boolean {
  const domain = email.toLowerCase().split('@')[1]
  return domain === 'gmail.com' || domain === 'googlemail.com'
}

export async function transferFileOwnership(
  fileId: string,
  targetEmail: string,
  currentUserEmail: string,
  emailMessage?: string,
  knownPermissions?: DrivePermission[],
): Promise<GrantResult> {
  const permissions = knownPermissions?.length ? knownPermissions : await listPermissions(fileId)
  const existing = permissions.find(
    (permission) =>
      permission.type === 'user' && permission.emailAddress?.toLowerCase() === targetEmail.toLowerCase(),
  )

  if (existing?.role === 'owner') {
    return { status: 'skipped', permission: existing, message: '目标账号已经是所有者' }
  }

  // Consumer Gmail requires the prospective owner to explicitly accept the transfer.
  if (isConsumerGmail(currentUserEmail) && isConsumerGmail(targetEmail)) {
    if (existing) {
      const params = new URLSearchParams({ supportsAllDrives: 'true', fields: 'id,type,role,emailAddress,displayName,pendingOwner' })
      const permission = await driveRequest<DrivePermission>(
        `/files/${encodeURIComponent(fileId)}/permissions/${encodeURIComponent(existing.id)}?${params}`,
        { method: 'PATCH', body: JSON.stringify({ role: 'writer', pendingOwner: true }) },
      )
      return { status: 'updated', permission, message: '已发起所有权转移，等待目标 Gmail 账号接受' }
    }

    const params = new URLSearchParams({
      supportsAllDrives: 'true',
      sendNotificationEmail: 'true',
      fields: 'id,type,role,emailAddress,displayName,pendingOwner',
    })
    if (emailMessage?.trim()) params.set('emailMessage', emailMessage.trim())
    const permission = await driveRequest<DrivePermission>(`/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
      method: 'POST',
      body: JSON.stringify({ type: 'user', role: 'writer', emailAddress: targetEmail, pendingOwner: true }),
    })
    return { status: 'created', permission, message: '已发起所有权转移，等待目标 Gmail 账号接受' }
  }

  // Workspace permits direct ownership transfer only within the same organization.
  if (existing) {
    const params = new URLSearchParams({
      supportsAllDrives: 'true',
      transferOwnership: 'true',
      fields: 'id,type,role,emailAddress,displayName',
    })
    const permission = await driveRequest<DrivePermission>(
      `/files/${encodeURIComponent(fileId)}/permissions/${encodeURIComponent(existing.id)}?${params}`,
      { method: 'PATCH', body: JSON.stringify({ role: 'owner' }) },
    )
    return { status: 'updated', permission, message: '所有权已转移；原所有者已变为编辑者' }
  }

  const params = new URLSearchParams({
    supportsAllDrives: 'true',
    transferOwnership: 'true',
    sendNotificationEmail: 'true',
    fields: 'id,type,role,emailAddress,displayName',
  })
  if (emailMessage?.trim()) params.set('emailMessage', emailMessage.trim())
  const permission = await driveRequest<DrivePermission>(`/files/${encodeURIComponent(fileId)}/permissions?${params}`, {
    method: 'POST',
    body: JSON.stringify({ type: 'user', role: 'owner', emailAddress: targetEmail }),
  })
  return { status: 'created', permission, message: '所有权已转移；原所有者已变为编辑者' }
}

export function sharingSummary(file: SheetFile): {
  anyone: number
  domain: number
  named: number
  risky: boolean
} {
  const permissions = file.permissions || []
  const anyone = permissions.filter((permission) => permission.type === 'anyone').length
  const domain = permissions.filter((permission) => permission.type === 'domain').length
  const named = permissions.filter((permission) => ['user', 'group'].includes(permission.type)).length
  return { anyone, domain, named, risky: anyone + domain > 0 }
}
