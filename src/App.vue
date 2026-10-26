<script setup lang="ts">
import { computed, nextTick, onMounted, ref, watch } from 'vue'
import {
  AppstoreOutlined,
  AuditOutlined,
  CheckCircleFilled,
  ClockCircleOutlined,
  CloudSyncOutlined,
  CopyOutlined,
  DeleteOutlined,
  DownloadOutlined,
  EyeOutlined,
  ExclamationCircleFilled,
  FilterOutlined,
  FileProtectOutlined,
  FileSearchOutlined,
  GoogleOutlined,
  HistoryOutlined,
  LinkOutlined,
  LockOutlined,
  LogoutOutlined,
  MailOutlined,
  ReloadOutlined,
  SafetyOutlined,
  SearchOutlined,
  SendOutlined,
  SettingOutlined,
  SortAscendingOutlined,
  StopOutlined,
  TeamOutlined,
} from '@ant-design/icons-vue'
import { message, Modal } from 'ant-design-vue'
import zhCN from 'ant-design-vue/es/locale/zh_CN'
import dayjs from 'dayjs'
import {
  addProtectedSheetEditors,
  connectGoogle,
  disconnectGoogle,
  enableFolderLimitedAccess,
  enableAnyoneLinkSharing,
  getDriveFile,
  grantPermission,
  hasChromeIdentity,
  isInheritedPermission,
  listDriveFiles,
  sharingSummary,
  tightenSharing,
  trashSpreadsheet,
  transferFileOwnership,
} from './lib/googleDrive'
import {
  countInvalidEmails,
  escapeCsvCell,
  mapWithConcurrency,
  parseEmails,
  parseGoogleFileIds,
  splitInput,
} from './lib/parsers'
import type {
  DrivePermission,
  GrantResult,
  OperationLog,
  PermissionRole,
  ProtectedSheetEditorResult,
  SheetFile,
  UserProfile,
} from './types'

type ResolvedTarget = { id: string; file?: SheetFile; error?: string }
type SharingFilter = 'all' | 'public' | 'domain' | 'restricted'
type OwnershipFilter = 'all' | 'mine' | 'others'
type ModifiedSort = 'newest' | 'oldest'
type GrantMappingMode = 'matrix' | 'paired' | 'public'
type FileTypeFilter = 'all' | 'spreadsheet' | 'document' | 'presentation' | 'folder' | 'pdf' | 'image' | 'other'
type AuditEmailFile = { id: string; name: string; url: string; role: string }
type AuditEmailGroup = {
  email: string
  displayName: string
  types: Array<'user' | 'group'>
  roles: string[]
  files: AuditEmailFile[]
}

const demoMode = !hasChromeIdentity()
const connected = ref(demoMode)
const connecting = ref(false)
const profile = ref<UserProfile>(demoMode ? { email: 'demo@gmail.com', id: 'demo' } : { email: '', id: '' })
const activeTab = ref('grant')
const scanLoading = ref(false)
const scanCount = ref(0)
const searchText = ref('')
const sharingFilter = ref<SharingFilter>('all')
const ownershipFilter = ref<OwnershipFilter>('all')
const modifiedSort = ref<ModifiedSort>('newest')
const fileTypeFilter = ref<FileTypeFilter>('all')
const currentPage = ref(1)
const pageSize = ref(10)
const selectedRowKeys = ref<string[]>([])
const tightenOpen = ref(false)
const tightenTypes = ref<Array<'anyone' | 'domain'>>(['anyone'])
const limitedAccessOnInherited = ref(false)
const tightenRunning = ref(false)
const openSharingOpen = ref(false)
const openSharingRole = ref<PermissionRole>('reader')
const openSharingConfirmed = ref(false)
const openSharingRunning = ref(false)
const trashOpen = ref(false)
const trashConfirmed = ref(false)
const trashRunning = ref(false)
const progress = ref({ visible: false, current: 0, total: 0, label: '' })

const linkInput = ref('')
const emailInput = ref('')
const grantMappingMode = ref<GrantMappingMode>('matrix')
const role = ref<PermissionRole | 'owner'>('reader')
const notifyEmail = ref(true)
const emailMessage = ref('')
const grantEnableLinkSharing = ref(false)
const grantLinkSharingRole = ref<PermissionRole>('reader')
const syncProtectedSheetEditors = ref(false)
const sheetKeywordInput = ref('')
const resolving = ref(false)
const granting = ref(false)
const ownershipOpen = ref(false)
const ownershipConfirmed = ref(false)
const resolvedTargets = ref<ResolvedTarget[]>([])
const logs = ref<OperationLog[]>([])

const auditInput = ref('')
const auditLoading = ref(false)
const auditResults = ref<ResolvedTarget[]>([])
const auditEmailSearch = ref('')

const GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet'
const GOOGLE_DOC_MIME = 'application/vnd.google-apps.document'
const GOOGLE_SLIDE_MIME = 'application/vnd.google-apps.presentation'
const GOOGLE_FOLDER_MIME = 'application/vnd.google-apps.folder'

const demoSheets: SheetFile[] = [
  {
    id: '1DemoQuarterlySalesSheet00001',
    name: 'Q4 销售数据汇总',
    mimeType: 'application/vnd.google-apps.spreadsheet',
    webViewLink: 'https://docs.google.com/spreadsheets/',
    modifiedTime: dayjs().subtract(8, 'minute').toISOString(),
    ownedByMe: true,
    parents: ['1DemoFinanceFolder00000001'],
    capabilities: { canShare: true, canTrash: true },
    owners: [{ displayName: '我', emailAddress: 'demo@gmail.com' }],
    permissions: [
      { id: 'owner-1', type: 'user', role: 'owner', emailAddress: 'demo@gmail.com' },
      { id: 'user-1', type: 'user', role: 'writer', emailAddress: 'ops@example.com' },
      { id: 'anyone-1', type: 'anyone', role: 'reader', allowFileDiscovery: false },
    ],
  },
  {
    id: '1DemoCampaignPlanSheet000002',
    name: '品牌活动排期',
    mimeType: 'application/vnd.google-apps.spreadsheet',
    webViewLink: 'https://docs.google.com/spreadsheets/',
    modifiedTime: dayjs().subtract(2, 'day').toISOString(),
    ownedByMe: true,
    parents: ['1DemoCampaignFolder0000001'],
    capabilities: { canShare: true, canTrash: true },
    owners: [{ displayName: '我', emailAddress: 'demo@gmail.com' }],
    permissions: [
      { id: 'owner-2', type: 'user', role: 'owner', emailAddress: 'demo@gmail.com' },
      {
        id: 'domain-1',
        type: 'domain',
        role: 'reader',
        domain: 'example.com',
        permissionDetails: [
          {
            permissionType: 'file',
            role: 'reader',
            inherited: true,
            inheritedFrom: '1DemoCampaignFolder0000001',
          },
        ],
      },
      { id: 'group-1', type: 'group', role: 'reader', emailAddress: 'marketing@example.com' },
    ],
  },
  {
    id: '1DemoHiringTrackerSheet00003',
    name: '招聘进度跟踪',
    mimeType: 'application/vnd.google-apps.spreadsheet',
    webViewLink: 'https://docs.google.com/spreadsheets/',
    modifiedTime: dayjs().subtract(5, 'day').toISOString(),
    ownedByMe: false,
    parents: ['1DemoPeopleFolder00000002'],
    capabilities: { canShare: false, canTrash: false },
    owners: [{ displayName: 'People Ops', emailAddress: 'people@example.com' }],
    permissions: [
      { id: 'owner-3', type: 'user', role: 'owner', emailAddress: 'people@example.com' },
      { id: 'user-3', type: 'user', role: 'reader', emailAddress: 'demo@gmail.com' },
    ],
  },
  {
    id: '1DemoProjectBriefDocument004',
    name: '新品项目说明书',
    mimeType: GOOGLE_DOC_MIME,
    webViewLink: 'https://docs.google.com/document/',
    modifiedTime: dayjs().subtract(12, 'hour').toISOString(),
    ownedByMe: true,
    capabilities: { canShare: true, canTrash: true },
    owners: [{ displayName: '我', emailAddress: 'demo@gmail.com' }],
    permissions: [
      { id: 'owner-4', type: 'user', role: 'owner', emailAddress: 'demo@gmail.com' },
      { id: 'user-4', type: 'user', role: 'reader', emailAddress: 'product@example.com' },
    ],
  },
  {
    id: '1DemoTeamAssetsFolder000005',
    name: '团队共享素材',
    mimeType: GOOGLE_FOLDER_MIME,
    webViewLink: 'https://drive.google.com/drive/folders/',
    modifiedTime: dayjs().subtract(4, 'day').toISOString(),
    ownedByMe: true,
    capabilities: { canShare: true, canTrash: true },
    owners: [{ displayName: '我', emailAddress: 'demo@gmail.com' }],
    permissions: [{ id: 'owner-5', type: 'user', role: 'owner', emailAddress: 'demo@gmail.com' }],
  },
]

const sheets = ref<SheetFile[]>(demoMode ? demoSheets : [])

const parsedFileIds = computed(() => parseGoogleFileIds(linkInput.value, grantMappingMode.value !== 'paired'))
const parsedAuditFileIds = computed(() => parseGoogleFileIds(auditInput.value))
const parsedEmails = computed(() => parseEmails(emailInput.value, grantMappingMode.value !== 'paired'))
const sheetKeywords = computed(() => [
  ...new Set(splitInput(sheetKeywordInput.value).map((keyword) => keyword.toLowerCase())),
])
const invalidEmailCount = computed(() => countInvalidEmails(emailInput.value))
const unrecognizedLinkCount = computed(() => Math.max(0, splitInput(linkInput.value).length - parsedFileIds.value.length))
const unrecognizedAuditLinkCount = computed(() =>
  Math.max(0, splitInput(auditInput.value).length - parsedAuditFileIds.value.length),
)
const validResolvedTargets = computed(() => resolvedTargets.value.filter((target) => target.file && !target.error))
const validAuditResults = computed(() => auditResults.value.filter((target) => target.file && !target.error))
const auditPublicCount = computed(
  () => validAuditResults.value.filter((target) => target.file && sharingSummary(target.file).anyone > 0).length,
)
const auditRestrictedCount = computed(
  () => validAuditResults.value.filter((target) => target.file && !sharingSummary(target.file).risky).length,
)
const currentAccountAuditFiles = computed<AuditEmailFile[]>(() =>
  validAuditResults.value.map((target) => ({
    id: target.file!.id,
    name: target.file!.name,
    url: spreadsheetLink(target.file!),
    role: '有效访问',
  })),
)
const auditEmailGroups = computed<AuditEmailGroup[]>(() => {
  const groups = new Map<
    string,
    {
      email: string
      displayName: string
      types: Set<'user' | 'group'>
      roles: Set<string>
      files: Map<string, AuditEmailFile>
    }
  >()

  for (const target of validAuditResults.value) {
    const file = target.file!
    for (const permission of file.permissions || []) {
      if (!['user', 'group'].includes(permission.type) || !permission.emailAddress) continue
      const email = permission.emailAddress.trim().toLowerCase()
      if (!email) continue
      const existing = groups.get(email) || {
        email,
        displayName: permission.displayName || '',
        types: new Set<'user' | 'group'>(),
        roles: new Set<string>(),
        files: new Map<string, AuditEmailFile>(),
      }
      if (!existing.displayName && permission.displayName) existing.displayName = permission.displayName
      existing.types.add(permission.type as 'user' | 'group')
      existing.roles.add(permission.role)
      existing.files.set(file.id, {
        id: file.id,
        name: file.name,
        url: spreadsheetLink(file),
        role: permission.role,
      })
      groups.set(email, existing)
    }
  }

  return [...groups.values()]
    .map((group) => ({
      email: group.email,
      displayName: group.displayName,
      types: [...group.types],
      roles: [...group.roles],
      files: [...group.files.values()],
    }))
    .sort((left, right) => right.files.length - left.files.length || left.email.localeCompare(right.email))
})
const filteredAuditEmailGroups = computed(() => {
  const query = auditEmailSearch.value.trim().toLowerCase()
  const currentEmail = profile.value.email.trim().toLowerCase()
  return auditEmailGroups.value.filter(
    (group) =>
      group.email !== currentEmail &&
      (!query || `${group.email} ${group.displayName}`.toLowerCase().includes(query)),
  )
})
const showCurrentAuditGroup = computed(() => {
  const query = auditEmailSearch.value.trim().toLowerCase()
  return !query || `当前登录邮箱 ${profile.value.email}`.toLowerCase().includes(query)
})
const publicGrantMode = computed(() => grantMappingMode.value === 'public')
const ownershipMode = computed(() => !publicGrantMode.value && role.value === 'owner')
const pairedMode = computed(() => grantMappingMode.value === 'paired')
const pairedCountMatches = computed(
  () => parsedFileIds.value.length > 0 && parsedFileIds.value.length === parsedEmails.value.length,
)
const pairedTasks = computed(() =>
  resolvedTargets.value.map((target, index) => ({ target, email: parsedEmails.value[index] || '' })),
)
const transferableTargets = computed(() =>
  validResolvedTargets.value.filter((target) => target.file && canTransferOwnership(target.file)),
)
const ownershipTargetEmail = computed(() => parsedEmails.value[0] || '')
function isConsumerOwnershipTarget(email: string): boolean {
  const currentDomain = profile.value.email.toLowerCase().split('@')[1]
  const targetDomain = email.toLowerCase().split('@')[1]
  return ['gmail.com', 'googlemail.com'].includes(currentDomain) && ['gmail.com', 'googlemail.com'].includes(targetDomain)
}
const consumerOwnershipFlow = computed(() => {
  return isConsumerOwnershipTarget(ownershipTargetEmail.value)
})
const permissionMatrixSize = computed(() => {
  if (publicGrantMode.value) return validResolvedTargets.value.length
  if (pairedMode.value) return resolvedTargets.value.length
  return ownershipMode.value ? validResolvedTargets.value.length : validResolvedTargets.value.length * parsedEmails.value.length
})
const protectedSheetTaskCount = computed(() =>
  !publicGrantMode.value && !ownershipMode.value && syncProtectedSheetEditors.value && sheetKeywords.value.length
    ? validResolvedTargets.value.filter((target) => target.file?.mimeType === GOOGLE_SHEET_MIME).length
    : 0,
)
const grantLinkSharingTaskCount = computed(() =>
  !publicGrantMode.value && !ownershipMode.value && grantEnableLinkSharing.value ? validResolvedTargets.value.length : 0,
)
const executionTaskCount = computed(
  () => permissionMatrixSize.value + protectedSheetTaskCount.value + grantLinkSharingTaskCount.value,
)
const executionTaskUnit = computed(() => {
  if (publicGrantMode.value) return '个公开链接权限'
  const extras: string[] = []
  if (grantLinkSharingTaskCount.value) extras.push(`${grantLinkSharingTaskCount.value} 个链接共享`)
  if (protectedSheetTaskCount.value) extras.push(`${protectedSheetTaskCount.value} 个 Sheet 检查`)
  return extras.length ? `项任务（含 ${extras.join('、')}）` : '项权限组合'
})

const filteredSheets = computed(() => {
  const query = searchText.value.trim().toLowerCase()
  return sheets.value
    .filter((sheet) => {
      if (fileTypeFilter.value !== 'all' && driveFileCategory(sheet) !== fileTypeFilter.value) return false
      if (query) {
        const owners = [
          ...(sheet.owners || []).flatMap((item) => [item.displayName || '', item.emailAddress || '']),
          ...(sheet.permissions || [])
            .filter((permission) => permission.role === 'owner')
            .flatMap((permission) => [permission.displayName || '', permission.emailAddress || '']),
          sheet.ownedByMe ? profile.value.email : '',
          sheet.driveId ? '共享云端硬盘 组织所有' : '',
        ]
        if (![sheet.name, sheet.id, ...owners].join(' ').toLowerCase().includes(query)) return false
      }

      const summary = sharingSummary(sheet)
      if (ownershipFilter.value === 'mine' && !isOwnedByCurrentUser(sheet)) return false
      if (ownershipFilter.value === 'others' && isOwnedByCurrentUser(sheet)) return false
      if (sharingFilter.value === 'public') return summary.anyone > 0
      if (sharingFilter.value === 'domain') return summary.domain > 0
      if (sharingFilter.value === 'restricted') return !summary.risky
      return true
    })
    .sort((left, right) => {
      const leftTime = left.modifiedTime ? Date.parse(left.modifiedTime) : Number.NaN
      const rightTime = right.modifiedTime ? Date.parse(right.modifiedTime) : Number.NaN
      if (Number.isNaN(leftTime)) return Number.isNaN(rightTime) ? 0 : 1
      if (Number.isNaN(rightTime)) return -1
      return modifiedSort.value === 'newest' ? rightTime - leftTime : leftTime - rightTime
    })
})

const hasActiveScanFilters = computed(
  () =>
    Boolean(searchText.value.trim()) ||
    sharingFilter.value !== 'all' ||
    ownershipFilter.value !== 'all' ||
    fileTypeFilter.value !== 'all' ||
    modifiedSort.value !== 'newest',
)

const tablePagination = computed(() => ({
  current: currentPage.value,
  pageSize: pageSize.value,
  pageSizeOptions: ['10', '20', '50', '100', '500', '1000'],
  showSizeChanger: true,
  showTotal: (total: number, range: [number, number]) => `第 ${range[0]}–${range[1]} 条，共 ${total} 条`,
  position: ['bottomRight'] as const,
}))

const publicSheetCount = computed(() => sheets.value.filter((sheet) => sharingSummary(sheet).anyone > 0).length)
const domainSheetCount = computed(() => sheets.value.filter((sheet) => sharingSummary(sheet).domain > 0).length)
const safeSheetCount = computed(() => sheets.value.filter((sheet) => !sharingSummary(sheet).risky).length)
const ownedSheetCount = computed(() => sheets.value.filter(isOwnedByCurrentUser).length)
const spreadsheetFileCount = computed(() => sheets.value.filter((file) => file.mimeType === GOOGLE_SHEET_MIME).length)
const selectedSheets = computed(() => sheets.value.filter((sheet) => selectedRowKeys.value.includes(sheet.id)))
const selectedExistingPublicCount = computed(
  () => selectedSheets.value.filter((sheet) => sharingSummary(sheet).anyone > 0).length,
)
const selectedTrashableCount = computed(
  () => selectedSheets.value.filter((sheet) => sheet.capabilities?.canTrash !== false).length,
)
const selectedRiskCount = computed(() =>
  selectedSheets.value.reduce(
    (sum, sheet) =>
      sum +
      (sheet.permissions || []).filter((permission) =>
        tightenTypes.value.includes(permission.type as 'anyone' | 'domain') && !isInheritedPermission(permission),
      ).length,
    0,
  ),
)
const selectedInheritedRiskCount = computed(() =>
  selectedSheets.value.reduce(
    (sum, sheet) =>
      sum +
      (sheet.permissions || []).filter(
        (permission) =>
          tightenTypes.value.includes(permission.type as 'anyone' | 'domain') && isInheritedPermission(permission),
      ).length,
    0,
  ),
)
const selectedParentFolderCount = computed(
  () => new Set(selectedSheets.value.flatMap((sheet) => sheet.parents || [])).size,
)

const columns = [
  { title: '文件', dataIndex: 'name', key: 'name', width: '38%' },
  { title: '所有者', key: 'owner', width: '23%' },
  { title: '共享状态', key: 'sharing', width: '24%' },
  { title: '最近修改', key: 'modified', width: '15%' },
]

watch(linkInput, () => {
  resolvedTargets.value = []
})

watch(grantMappingMode, (value, previous) => {
  if (value === 'paired' || previous === 'paired') resolvedTargets.value = []
  ownershipConfirmed.value = false
  if (value === 'public') {
    role.value = 'reader'
    grantLinkSharingRole.value = 'writer'
    grantEnableLinkSharing.value = false
    syncProtectedSheetEditors.value = false
  }
})

watch(auditInput, () => {
  auditResults.value = []
  auditEmailSearch.value = ''
})

watch([searchText, sharingFilter, ownershipFilter, fileTypeFilter, modifiedSort], () => {
  currentPage.value = 1
})

watch(role, (value) => {
  if (value === 'owner') {
    notifyEmail.value = true
    grantEnableLinkSharing.value = false
    syncProtectedSheetEditors.value = false
  }
})

watch(
  logs,
  (value) => {
    try {
      localStorage.setItem('sheet-guardian-logs', JSON.stringify(value.slice(0, 500)))
    } catch {
      // Permission operations must not fail merely because local history cannot be persisted.
    }
  },
  { deep: true },
)

function addLog(log: Omit<OperationLog, 'id' | 'time'>) {
  logs.value.unshift({ ...log, id: crypto.randomUUID(), time: new Date().toISOString() })
  if (logs.value.length > 500) logs.value.length = 500
}

async function handleConnect() {
  if (demoMode) {
    message.info('当前为网页演示模式；构建后在 Chrome 扩展中会连接真实 Google 账号')
    return
  }
  connecting.value = true
  try {
    profile.value = await connectGoogle(true)
    connected.value = true
    message.success(`已连接 ${profile.value.email || '当前 Chrome Google 账号'}`)
    await scanSheets()
  } catch (error) {
    message.error(error instanceof Error ? error.message : 'Google 账号连接失败')
  } finally {
    connecting.value = false
  }
}

async function handleDisconnect() {
  if (demoMode) return message.info('演示模式不需要退出')
  await disconnectGoogle()
  connected.value = false
  profile.value = { email: '', id: '' }
  sheets.value = []
  selectedRowKeys.value = []
  message.success('已退出当前授权会话')
}

async function scanSheets() {
  if (!connected.value) return message.warning('请先连接 Google 账号')
  scanLoading.value = true
  scanCount.value = 0
  try {
    if (demoMode) {
      await new Promise((resolve) => setTimeout(resolve, 450))
      sheets.value = [...demoSheets]
      scanCount.value = demoSheets.length
    } else {
      sheets.value = await listDriveFiles((count) => (scanCount.value = count))
    }
    selectedRowKeys.value = selectedRowKeys.value.filter((id) => sheets.value.some((sheet) => sheet.id === id))
    currentPage.value = 1
    message.success(`扫描完成，共找到 ${sheets.value.length} 个 Google Drive 文件`)
  } catch (error) {
    message.error(error instanceof Error ? error.message : '扫描失败')
  } finally {
    scanLoading.value = false
  }
}

function handleTableChange(pagination: { current?: number; pageSize?: number }) {
  const nextSize = pagination.pageSize || 10
  currentPage.value = nextSize !== pageSize.value ? 1 : pagination.current || 1
  pageSize.value = nextSize
}

function clearSelectedSheets() {
  selectedRowKeys.value = []
  message.success('已取消全部选择')
}

function isOwnedByCurrentUser(file: SheetFile): boolean {
  if (file.driveId) return false
  if (file.ownedByMe) return true
  const currentEmail = profile.value.email.trim().toLowerCase()
  if (!currentEmail) return false
  return Boolean(
    [
      ...(file.owners || []).map((owner) => owner.emailAddress || ''),
      ...(file.permissions || [])
        .filter((permission) => permission.role === 'owner')
        .map((permission) => permission.emailAddress || ''),
    ].some((email) => email.toLowerCase() === currentEmail),
  )
}

function selectAllOwnedSheets() {
  const ownedIds = sheets.value
    .filter((sheet) => isOwnedByCurrentUser(sheet) && sheet.capabilities?.canShare !== false)
    .map((sheet) => sheet.id)
  selectedRowKeys.value = ownedIds
  ownershipFilter.value = 'mine'
  message.success(`已选择当前账号拥有的 ${ownedIds.length} 个可授权文件`)
}

async function sendSelectedToGrant() {
  const targets = selectedSheets.value.filter((sheet) => sheet.capabilities?.canShare !== false)
  if (!targets.length) return message.warning('请选择至少一个当前账号可授权的文件')
  linkInput.value = targets.map((sheet) => sheet.id).join('\n')
  await nextTick()
  resolvedTargets.value = targets.map((file) => ({ id: file.id, file }))
  emailInput.value = ''
  emailMessage.value = ''
  grantEnableLinkSharing.value = false
  grantLinkSharingRole.value = publicGrantMode.value ? 'writer' : 'reader'
  syncProtectedSheetEditors.value = false
  sheetKeywordInput.value = ''
  role.value = 'reader'
  notifyEmail.value = true
  activeTab.value = 'grant'
  message.success(
    publicGrantMode.value
      ? `已带入 ${targets.length} 个文件，可直接开启公开链接权限`
      : `已带入 ${targets.length} 个文件，请填写需要添加的邮箱`,
  )
}

function resetScanFilters() {
  searchText.value = ''
  sharingFilter.value = 'all'
  ownershipFilter.value = 'all'
  fileTypeFilter.value = 'all'
  modifiedSort.value = 'newest'
}

async function resolveLinks() {
  if (!connected.value) return message.warning('请先连接 Google 账号')
  if (!parsedFileIds.value.length) return message.warning('请至少输入一个有效的 Google Drive 文件链接或 ID')
  resolving.value = true
  try {
    resolvedTargets.value = await mapWithConcurrency(parsedFileIds.value, 4, async (id) => {
      if (demoMode) {
        const existing = demoSheets.find((sheet) => sheet.id === id)
        return existing
          ? { id, file: existing }
          : {
              id,
              file: {
                ...demoSheets[0],
                id,
                name: `演示文件 · ${id.slice(0, 8)}`,
                permissions: [{ id: `owner-${id}`, type: 'user', role: 'owner', emailAddress: profile.value.email }],
              },
            }
      }
      try {
        return { id, file: await getDriveFile(id) }
      } catch (error) {
        return { id, error: error instanceof Error ? error.message : '无法读取此文件' }
      }
    })
    const failures = resolvedTargets.value.filter((target) => target.error).length
    if (failures) message.warning(`已解析 ${validResolvedTargets.value.length} 个文件，${failures} 个链接不可用`)
    else message.success(`已确认 ${validResolvedTargets.value.length} 个目标文件`)
  } finally {
    resolving.value = false
  }
}

async function inspectLinks() {
  if (!connected.value) return message.warning('请先连接 Google 账号')
  if (!parsedAuditFileIds.value.length) return message.warning('请至少输入一个有效的 Google Drive 文件链接或 ID')
  auditLoading.value = true
  try {
    auditResults.value = await mapWithConcurrency(parsedAuditFileIds.value, 4, async (id) => {
      if (demoMode) {
        await new Promise((resolve) => setTimeout(resolve, 90))
        const existing = demoSheets.find((sheet) => sheet.id === id)
        return existing
          ? { id, file: existing }
          : {
              id,
              file: {
                ...demoSheets[0],
                id,
                name: `演示文件 · ${id.slice(0, 8)}`,
                permissions: [
                  { id: `owner-${id}`, type: 'user', role: 'owner', emailAddress: profile.value.email },
                  { id: `reader-${id}`, type: 'user', role: 'reader', emailAddress: 'viewer@example.com' },
                ],
              },
            }
      }
      try {
        return { id, file: await getDriveFile(id) }
      } catch (error) {
        return { id, error: error instanceof Error ? error.message : '无法读取此文件' }
      }
    })
    const failures = auditResults.value.filter((target) => target.error).length
    if (failures) message.warning(`已读取 ${validAuditResults.value.length} 个文件，${failures} 个链接无法访问`)
    else message.success(`已读取 ${validAuditResults.value.length} 个文件的权限信息`)
  } finally {
    auditLoading.value = false
  }
}

function clearAudit() {
  auditInput.value = ''
  auditResults.value = []
  auditEmailSearch.value = ''
}

function spreadsheetLink(file: SheetFile): string {
  const fallback = `https://drive.google.com/open?id=${encodeURIComponent(file.id)}`
  if (!file.webViewLink) return fallback

  try {
    const url = new URL(file.webViewLink)
    if (url.protocol === 'https:' && ['docs.google.com', 'drive.google.com'].includes(url.hostname)) return url.href
  } catch {
    // Fall through to the trusted Drive URL.
  }

  return fallback
}

function driveFileCategory(file: SheetFile): Exclude<FileTypeFilter, 'all'> {
  if (file.mimeType === GOOGLE_SHEET_MIME) return 'spreadsheet'
  if (file.mimeType === GOOGLE_DOC_MIME) return 'document'
  if (file.mimeType === GOOGLE_SLIDE_MIME) return 'presentation'
  if (file.mimeType === GOOGLE_FOLDER_MIME) return 'folder'
  if (file.mimeType === 'application/pdf') return 'pdf'
  if (file.mimeType.startsWith('image/')) return 'image'
  return 'other'
}

function driveFileTypeLabel(file: SheetFile): string {
  return {
    spreadsheet: 'Google 表格',
    document: 'Google 文档',
    presentation: 'Google 幻灯片',
    folder: '文件夹',
    pdf: 'PDF',
    image: '图片',
    other: '其他文件',
  }[driveFileCategory(file)]
}

function driveFileIconText(file: SheetFile): string {
  return {
    spreadsheet: 'S',
    document: 'D',
    presentation: 'P',
    folder: 'F',
    pdf: 'PDF',
    image: 'IMG',
    other: 'FILE',
  }[driveFileCategory(file)]
}

function ownerText(file: SheetFile): string {
  if (file.driveId) return '共享云端硬盘（归组织所有）'
  const owners = file.owners?.length
    ? file.owners
    : (file.permissions || [])
        .filter((permission) => permission.role === 'owner')
        .map((permission) => ({ displayName: permission.displayName, emailAddress: permission.emailAddress }))
  if (!owners.length && file.ownedByMe) return profile.value.email || '当前账号'
  return owners.map((owner) => owner.emailAddress || owner.displayName || '未知所有者').join('、') || '未知所有者'
}

function namedPermissions(file: SheetFile): DrivePermission[] {
  return (file.permissions || []).filter((permission) => ['user', 'group'].includes(permission.type))
}

function permissionTarget(permission: DrivePermission): string {
  if (permission.type === 'anyone') return permission.allowFileDiscovery ? '互联网公开（可搜索）' : '知道链接的任何人'
  if (permission.type === 'domain') return permission.domain ? `${permission.domain} 域` : 'Workspace 域'
  return permission.emailAddress || permission.displayName || (permission.type === 'group' ? '未显示的群组' : '未显示的账号')
}

function permissionRoleLabel(value: string): string {
  return {
    owner: '所有者',
    organizer: '管理者',
    fileOrganizer: '内容管理员',
    writer: '可编辑',
    commenter: '可评论',
    reader: '可查看',
  }[value] || value
}

function permissionTypeLabel(value: DrivePermission['type']): string {
  return { user: '账号', group: '群组', domain: '域', anyone: '公开' }[value] || value
}

async function copyAuditResults() {
  if (!auditResults.value.length) return message.info('暂无可复制的查询结果')
  const lines = auditResults.value.flatMap((target, index) => {
    if (!target.file) return [`${index + 1}. ${target.id}`, `查询失败：${target.error || '无法读取'}`, '']
    const file = target.file
    const summary = sharingSummary(file)
    const permissions = (file.permissions || []).map(
      (permission) =>
        `- ${permissionTypeLabel(permission.type)}｜${permissionTarget(permission)}｜${permissionRoleLabel(permission.role)}${
          permission.pendingOwner ? '（待接受所有权）' : ''
        }`,
    )
    return [
      `${index + 1}. ${file.name}`,
      `文件 ID：${file.id}`,
      `所有者：${ownerText(file)}`,
      `共享范围：${summary.anyone ? '公开链接' : summary.domain ? '域内共享' : '受限'}`,
      `当前账号管理权限：${file.capabilities?.canShare === false ? '不可管理' : '可管理共享'}`,
      '权限明细：',
      ...(permissions.length ? permissions : ['- Google 未返回权限明细']),
      '',
    ]
  })
  try {
    await navigator.clipboard.writeText(lines.join('\n'))
    message.success('权限查询结果已复制')
  } catch {
    message.error('复制失败，请检查浏览器剪贴板权限')
  }
}

async function copyAuditEmailLinks(label: string, files: AuditEmailFile[]) {
  if (!files.length) return message.info(`${label} 暂无可复制的文件链接`)
  const links = [...new Set(files.map((file) => file.url))]
  try {
    await navigator.clipboard.writeText(links.join('\n'))
    message.success(`已复制 ${label} 的 ${links.length} 个文件链接`)
  } catch {
    message.error('复制失败，请检查浏览器剪贴板权限')
  }
}

function canTransferOwnership(file: SheetFile): boolean {
  if (file.driveId || file.capabilities?.canShare === false) return false
  if (file.ownedByMe) return true
  const currentEmail = profile.value.email.toLowerCase()
  return Boolean(
    currentEmail &&
      [
        ...(file.owners || []).map((owner) => owner.emailAddress || ''),
        ...(file.permissions || [])
          .filter((permission) => permission.role === 'owner')
          .map((permission) => permission.emailAddress || ''),
      ].some((email) => email.toLowerCase() === currentEmail),
  )
}

function requestBatchExecution() {
  if (!resolvedTargets.value.length || !validResolvedTargets.value.length) return message.warning('请先解析并确认文件链接')
  if (!publicGrantMode.value && !parsedEmails.value.length) return message.warning('请至少输入一个有效邮箱')
  if (!publicGrantMode.value && pairedMode.value && !pairedCountMatches.value) {
    return message.warning(`一一对应要求文件与邮箱数量相同：当前 ${parsedFileIds.value.length} 个文件、${parsedEmails.value.length} 个邮箱`)
  }
  if (!ownershipMode.value && syncProtectedSheetEditors.value && !sheetKeywords.value.length) {
    return message.warning('已勾选 Sheet 保护权限，请至少填写一个 Sheet 关键词')
  }
  if (ownershipMode.value) {
    if (!pairedMode.value && parsedEmails.value.length !== 1) return message.warning('组合模式转移所有者时只能填写一个目标邮箱')
    if (!transferableTargets.value.length) return message.warning('这些文件不属于当前账号，或位于共享云端硬盘，无法转移')
    ownershipConfirmed.value = false
    ownershipOpen.value = true
    return
  }
  void executeGrant()
}

async function confirmOwnershipTransfer() {
  ownershipOpen.value = false
  await executeGrant()
}

async function executeGrant() {
  if (!resolvedTargets.value.length || !validResolvedTargets.value.length) return message.warning('请先解析并确认文件链接')
  if (!publicGrantMode.value && !parsedEmails.value.length) return message.warning('请至少输入一个有效邮箱')
  if (!publicGrantMode.value && pairedMode.value && !pairedCountMatches.value) {
    return message.warning(`一一对应要求文件与邮箱数量相同：当前 ${parsedFileIds.value.length} 个文件、${parsedEmails.value.length} 个邮箱`)
  }
  if (!ownershipMode.value && syncProtectedSheetEditors.value && !sheetKeywords.value.length) {
    return message.warning('已勾选 Sheet 保护权限，请至少填写一个 Sheet 关键词')
  }
  if (ownershipMode.value && !pairedMode.value && parsedEmails.value.length !== 1) {
    return message.warning('组合模式转移所有者时只能填写一个目标邮箱')
  }

  granting.value = true
  progress.value = {
    visible: true,
    current: 0,
    total: executionTaskCount.value,
    label: publicGrantMode.value
      ? `正在开启公开链接${roleLabel(grantLinkSharingRole.value)}权限`
      : ownershipMode.value
      ? '正在批量转移所有权'
      : grantEnableLinkSharing.value || syncProtectedSheetEditors.value
        ? '正在设置邮箱与附加共享权限'
        : '正在批量设置权限',
  }

  const targetsToExecute = pairedMode.value ? resolvedTargets.value : validResolvedTargets.value
  await mapWithConcurrency(targetsToExecute, 2, async (target, targetIndex) => {
    const targetEmails = pairedMode.value ? [parsedEmails.value[targetIndex]].filter(Boolean) : parsedEmails.value
    if (!target.file) {
      const targetEmail = targetEmails[0] || ''
      addLog({
        action: ownershipMode.value ? '转移所有权' : '授权',
        status: 'failed',
        fileId: target.id,
        fileName: '无法读取的文件',
        target: targetEmail,
        detail: target.error || '无法读取此文件',
      })
      progress.value.current++
      return
    }
    const file = target.file
    const permissionCache = [...(file.permissions || [])]

    if (publicGrantMode.value) {
      try {
        if (file.capabilities?.canShare === false) throw new Error('当前账号无权修改此文件的共享设置')
        const result = await applyAnyoneLinkSharing(file, grantLinkSharingRole.value)
        addLog({
          action: '开启共享',
          status: result.status === 'skipped' ? 'skipped' : 'success',
          fileId: file.id,
          fileName: file.name,
          target: '知道链接的任何人',
          detail: result.message,
        })
      } catch (error) {
        addLog({
          action: '开启共享',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: '知道链接的任何人',
          detail: error instanceof Error ? error.message : '开启公开链接权限失败',
        })
      } finally {
        progress.value.current++
      }
      return
    }

    if (ownershipMode.value) {
      const email = targetEmails[0]
      if (!canTransferOwnership(file)) {
        addLog({
          action: '转移所有权',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: file.driveId ? '共享云端硬盘中的文件归组织所有，不支持转移' : '当前账号不是所有者或无权转移',
        })
        progress.value.current++
        return
      }

      try {
        if (demoMode) await new Promise((resolve) => setTimeout(resolve, 120))
        const result = demoMode
          ? {
              status: 'updated' as const,
              message: isConsumerOwnershipTarget(email)
                ? '已发起所有权转移，等待目标 Gmail 账号接受'
                : '所有权已转移；原所有者已变为编辑者',
            }
          : await transferFileOwnership(file.id, email, profile.value.email, emailMessage.value, permissionCache)

        if (result.permission) {
          const existingIndex = permissionCache.findIndex((permission) => permission.id === result.permission?.id)
          if (existingIndex >= 0) permissionCache[existingIndex] = result.permission
          else permissionCache.push(result.permission)
          file.permissions = permissionCache
        }
        addLog({
          action: '转移所有权',
          status: result.status === 'skipped' ? 'skipped' : 'success',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: result.message,
        })
        if (result.message.startsWith('所有权已转移')) {
          file.ownedByMe = false
          file.owners = [{ displayName: email, emailAddress: email }]
          file.permissions = permissionCache.map((permission) => {
            if (permission.emailAddress?.toLowerCase() === profile.value.email.toLowerCase()) {
              return { ...permission, role: 'writer' }
            }
            if (permission.emailAddress?.toLowerCase() === email.toLowerCase()) {
              return { ...permission, role: 'owner' }
            }
            return permission
          })
        }
      } catch (error) {
        addLog({
          action: '转移所有权',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: error instanceof Error ? error.message : '所有权转移失败',
        })
      } finally {
        progress.value.current++
      }
      return
    }

    const protectedEditorEmails: string[] = []
    for (const email of targetEmails) {
      if (file.capabilities?.canShare === false) {
        addLog({
          action: '授权',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: '当前账号无权修改此文件的共享设置',
        })
        progress.value.current++
        continue
      }
      try {
        if (demoMode) await new Promise((resolve) => setTimeout(resolve, 90))
        const result = demoMode
          ? { status: 'created' as const, message: '权限已添加' }
          : await grantPermission(
              file.id,
              email,
              role.value as PermissionRole,
              notifyEmail.value,
              emailMessage.value,
              permissionCache,
            )
        if (result.permission && !permissionCache.some((permission) => permission.id === result.permission?.id)) {
          permissionCache.push(result.permission)
        }
        addLog({
          action: '授权',
          status: result.status === 'skipped' ? 'skipped' : 'success',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: result.message,
        })
        protectedEditorEmails.push(email)
      } catch (error) {
        addLog({
          action: '授权',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: email,
          detail: error instanceof Error ? error.message : '授权失败',
        })
      } finally {
        progress.value.current++
      }
    }
    file.permissions = permissionCache

    if (grantEnableLinkSharing.value) {
      try {
        if (file.capabilities?.canShare === false) throw new Error('当前账号无权修改此文件的共享设置')
        const result = await applyAnyoneLinkSharing(file, grantLinkSharingRole.value)
        addLog({
          action: '开启共享',
          status: result.status === 'skipped' ? 'skipped' : 'success',
          fileId: file.id,
          fileName: file.name,
          target: '知道链接的任何人',
          detail: result.message,
        })
      } catch (error) {
        addLog({
          action: '开启共享',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: '知道链接的任何人',
          detail: error instanceof Error ? error.message : '开启链接共享失败',
        })
      } finally {
        progress.value.current++
      }
    }

    if (syncProtectedSheetEditors.value && file.mimeType === GOOGLE_SHEET_MIME) {
      try {
        let result: ProtectedSheetEditorResult
        if (!protectedEditorEmails.length) {
          result = emptyProtectedSheetResult()
        } else if (demoMode) {
          await new Promise((resolve) => setTimeout(resolve, 110))
          result = demoProtectedSheetResult(file, sheetKeywords.value)
        } else {
          result = await addProtectedSheetEditors(file.id, sheetKeywords.value, protectedEditorEmails)
        }
        addLog({
          action: 'Sheet 保护权限',
          status: result.updatedRanges > 0 ? 'success' : 'skipped',
          fileId: file.id,
          fileName: file.name,
          target: protectedEditorEmails.join(', ') || parsedEmails.value.join(', '),
          detail: protectedSheetResultDetail(result, protectedEditorEmails.length > 0),
        })
      } catch (error) {
        addLog({
          action: 'Sheet 保护权限',
          status: 'failed',
          fileId: file.id,
          fileName: file.name,
          target: protectedEditorEmails.join(', ') || parsedEmails.value.join(', '),
          detail: error instanceof Error ? error.message : 'Sheet 保护权限更新失败',
        })
      } finally {
        progress.value.current++
      }
    }
  })

  granting.value = false
  message.success(
    publicGrantMode.value
      ? '公开链接权限设置完成，请查看操作记录'
      : ownershipMode.value
        ? '批量所有权任务已完成，请查看操作记录'
        : '批量授权任务已完成，可在操作记录中查看结果',
  )
  activeTab.value = 'history'
}

function emptyProtectedSheetResult(): ProtectedSheetEditorResult {
  return {
    matchedSheets: 0,
    sheetsWithProtection: 0,
    eligibleRanges: 0,
    updatedRanges: 0,
    alreadyPresentRanges: 0,
    skippedWarningOnly: 0,
    skippedUneditable: 0,
  }
}

function demoProtectedSheetResult(file: SheetFile, keywords: string[]): ProtectedSheetEditorResult {
  const demoTabs: Record<string, Array<{ title: string; protected: boolean }>> = {
    '1DemoQuarterlySalesSheet00001': [
      { title: '财务汇总', protected: true },
      { title: '收入明细', protected: true },
      { title: '原始数据', protected: false },
    ],
    '1DemoCampaignPlanSheet000002': [
      { title: '活动排期', protected: true },
      { title: '审批记录', protected: true },
    ],
  }
  const matched = (demoTabs[file.id] || []).filter((sheet) =>
    keywords.some((keyword) => sheet.title.toLowerCase().includes(keyword)),
  )
  const protectedCount = matched.filter((sheet) => sheet.protected).length
  return {
    ...emptyProtectedSheetResult(),
    matchedSheets: matched.length,
    sheetsWithProtection: protectedCount,
    eligibleRanges: protectedCount,
    updatedRanges: protectedCount,
  }
}

function protectedSheetResultDetail(result: ProtectedSheetEditorResult, hasGrantedEmails: boolean): string {
  if (!hasGrantedEmails) return '文件共享授权均未成功，因此未修改 Sheet 保护权限'
  if (!result.matchedSheets) return '没有 Sheet 标题包含指定关键词，已跳过'
  if (!result.sheetsWithProtection) {
    return `匹配 ${result.matchedSheets} 个 Sheet，但均未设置保护权限，已跳过`
  }

  const details = [
    `匹配 ${result.matchedSheets} 个 Sheet`,
    `${result.sheetsWithProtection} 个存在保护规则`,
    `已更新 ${result.updatedRanges} 条`,
  ]
  if (result.alreadyPresentRanges) details.push(`${result.alreadyPresentRanges} 条已包含这些邮箱`)
  if (result.skippedWarningOnly) details.push(`${result.skippedWarningOnly} 条仅警告规则已跳过`)
  if (result.skippedUneditable) details.push(`${result.skippedUneditable} 条无权修改已跳过`)
  return details.join('；')
}

async function applyAnyoneLinkSharing(file: SheetFile, linkRole: PermissionRole): Promise<GrantResult> {
  const permissionCache = [...(file.permissions || [])]
  let result: GrantResult
  if (demoMode) {
    await new Promise((resolve) => setTimeout(resolve, 120))
    const existing = permissionCache.find((permission) => permission.type === 'anyone')
    if (existing?.role === linkRole) {
      result = { status: 'skipped', permission: existing, message: `链接共享已经是 ${linkRole}` }
    } else if (existing && isInheritedPermission(existing)) {
      result = {
        status: 'skipped',
        permission: existing,
        message: `链接共享继承自上级位置，已保留现有 ${existing.role} 权限`,
      }
    } else if (existing) {
      result = {
        status: 'updated',
        permission: { ...existing, role: linkRole },
        message: `链接共享已更新为 ${linkRole}`,
      }
    } else {
      result = {
        status: 'created',
        permission: {
          id: `anyone-${file.id}`,
          type: 'anyone',
          role: linkRole,
          allowFileDiscovery: false,
        },
        message: `已开启知道链接的任何人（${linkRole}）`,
      }
    }
  } else {
    result = await enableAnyoneLinkSharing(file.id, linkRole, permissionCache)
  }

  if (result.permission) {
    const existingIndex = permissionCache.findIndex((permission) => permission.id === result.permission?.id)
    if (existingIndex >= 0) permissionCache[existingIndex] = result.permission
    else permissionCache.push(result.permission)
    file.permissions = permissionCache
  }
  return result
}

function openEnableSharingModal() {
  if (!selectedRowKeys.value.length) return message.warning('请先选择至少一个文件')
  openSharingRole.value = 'reader'
  openSharingConfirmed.value = false
  openSharingOpen.value = true
}

async function executeEnableSharing() {
  if (!openSharingConfirmed.value) return message.warning('请先确认公开链接共享风险')
  openSharingRunning.value = true
  progress.value = {
    visible: true,
    current: 0,
    total: selectedSheets.value.length,
    label: '正在批量开启链接共享',
  }

  await mapWithConcurrency(selectedSheets.value, 3, async (file) => {
    if (file.capabilities?.canShare === false) {
      addLog({
        action: '开启共享',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        target: '知道链接的任何人',
        detail: '当前账号无权修改此文件的共享设置',
      })
      progress.value.current++
      return
    }

    try {
      const result = await applyAnyoneLinkSharing(file, openSharingRole.value)
      addLog({
        action: '开启共享',
        status: result.status === 'skipped' ? 'skipped' : 'success',
        fileId: file.id,
        fileName: file.name,
        target: '知道链接的任何人',
        detail: result.message,
      })
    } catch (error) {
      addLog({
        action: '开启共享',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        target: '知道链接的任何人',
        detail: error instanceof Error ? error.message : '开启链接共享失败',
      })
    } finally {
      progress.value.current++
    }
  })

  openSharingRunning.value = false
  openSharingOpen.value = false
  message.success('批量开启共享任务已完成，请在操作记录中查看结果')
}

function openTrashModal() {
  if (!selectedRowKeys.value.length) return message.warning('请先选择至少一个文件')
  trashConfirmed.value = false
  trashOpen.value = true
}

async function executeTrash() {
  if (!trashConfirmed.value) return message.warning('请先确认移入回收站')
  const targets = [...selectedSheets.value]
  const trashedIds = new Set<string>()
  let failedCount = 0
  trashRunning.value = true
  progress.value = {
    visible: true,
    current: 0,
    total: targets.length,
    label: '正在批量移入 Google Drive 回收站',
  }

  await mapWithConcurrency(targets, 3, async (file) => {
    if (file.capabilities?.canTrash === false) {
      failedCount++
      addLog({
        action: '移至回收站',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        target: 'Google Drive 回收站',
        detail: '当前账号无权将此文件移入回收站',
      })
      progress.value.current++
      return
    }

    try {
      if (demoMode) await new Promise((resolve) => setTimeout(resolve, 120))
      else await trashSpreadsheet(file.id)
      trashedIds.add(file.id)
      addLog({
        action: '移至回收站',
        status: 'success',
        fileId: file.id,
        fileName: file.name,
        target: 'Google Drive 回收站',
        detail: '已移入回收站，可在 Google Drive 回收站中恢复',
      })
    } catch (error) {
      failedCount++
      addLog({
        action: '移至回收站',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        target: 'Google Drive 回收站',
        detail: error instanceof Error ? error.message : '移入回收站失败',
      })
    } finally {
      progress.value.current++
    }
  })

  sheets.value = sheets.value.filter((file) => !trashedIds.has(file.id))
  selectedRowKeys.value = selectedRowKeys.value.filter((id) => !trashedIds.has(id))
  const maxPage = Math.max(1, Math.ceil(filteredSheets.value.length / pageSize.value))
  currentPage.value = Math.min(currentPage.value, maxPage)
  trashRunning.value = false
  trashOpen.value = false

  if (trashedIds.size) {
    message.success(`已将 ${trashedIds.size} 个文件移入回收站${failedCount ? `，${failedCount} 个失败` : ''}`)
  } else {
    message.warning('没有文件成功移入回收站，请查看操作记录')
  }
}

function openTightenModal() {
  if (!selectedRowKeys.value.length) return message.warning('请先选择至少一个文件')
  tightenTypes.value = ['anyone']
  limitedAccessOnInherited.value = false
  tightenOpen.value = true
}

async function executeTighten() {
  if (!tightenTypes.value.length) return message.warning('请选择要关闭的共享类型')
  tightenRunning.value = true
  progress.value = { visible: true, current: 0, total: selectedSheets.value.length, label: '正在收紧共享范围' }
  const limitedAccessTasks = new Map<string, Promise<{ id: string; name: string }>>()

  await mapWithConcurrency(selectedSheets.value, 3, async (file) => {
    if (file.capabilities?.canShare === false) {
      addLog({
        action: '收紧共享',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        detail: '当前账号无权修改此文件的共享设置',
      })
      progress.value.current++
      return
    }
    try {
      const demoCandidates = (file.permissions || []).filter((permission) =>
        tightenTypes.value.includes(permission.type as 'anyone' | 'domain'),
      )
      const result = demoMode
        ? {
            removed: demoCandidates.filter((permission) => !isInheritedPermission(permission)),
            skippedInherited: demoCandidates.filter(isInheritedPermission),
          }
        : await tightenSharing(file.id, tightenTypes.value, file.permissions)
      if (demoMode) await new Promise((resolve) => setTimeout(resolve, 150))
      const removedIds = new Set(result.removed.map((permission) => permission.id))
      file.permissions = (file.permissions || []).filter((permission) => !removedIds.has(permission.id))

      let limitedAccessEnabled = 0
      const limitedAccessErrors: string[] = []
      if (result.skippedInherited.length && limitedAccessOnInherited.value) {
        const parentIds = [...new Set(file.parents || [])]
        if (!parentIds.length) {
          limitedAccessErrors.push('Google Drive 未返回直属父文件夹，无法启用有限访问')
        }

        const folderResults = await Promise.all(
          parentIds.map(async (folderId) => {
            try {
              let task = limitedAccessTasks.get(folderId)
              if (!task) {
                task = demoMode
                  ? Promise.resolve({ id: folderId, name: `演示文件夹 · ${folderId.slice(0, 8)}` })
                  : enableFolderLimitedAccess(folderId).then((folder) => ({ id: folder.id, name: folder.name }))
                limitedAccessTasks.set(folderId, task)
              }
              const folder = await task
              return { ok: true as const, folder }
            } catch (error) {
              return {
                ok: false as const,
                error: error instanceof Error ? error.message : `文件夹 ${folderId} 启用有限访问失败`,
              }
            }
          }),
        )
        limitedAccessEnabled = folderResults.filter((item) => item.ok).length
        limitedAccessErrors.push(...folderResults.filter((item) => !item.ok).map((item) => item.error))
      }

      const handledInherited = limitedAccessOnInherited.value && limitedAccessEnabled > 0
      const operationStatus =
        result.removed.length || handledInherited
          ? 'success'
          : limitedAccessOnInherited.value && limitedAccessErrors.length
            ? 'failed'
            : result.skippedInherited.length
              ? 'skipped'
              : 'success'
      addLog({
        action: '收紧共享',
        status: operationStatus,
        fileId: file.id,
        fileName: file.name,
        detail: `已关闭 ${result.removed.length} 项直接链接/域共享；保留所有明确添加的账号${
          result.skippedInherited.length
            ? limitedAccessOnInherited.value
              ? `；检测到 ${result.skippedInherited.length} 项继承权限，已为 ${limitedAccessEnabled} 个直属文件夹启用 limited access${
                  limitedAccessErrors.length ? `；${limitedAccessErrors.join('；')}` : ''
                }`
              : `；已跳过 ${result.skippedInherited.length} 项继承权限，请到上级文件夹或共享云端硬盘调整`
            : ''
        }`,
      })
    } catch (error) {
      addLog({
        action: '收紧共享',
        status: 'failed',
        fileId: file.id,
        fileName: file.name,
        detail: error instanceof Error ? error.message : '操作失败',
      })
    } finally {
      progress.value.current++
    }
  })

  tightenRunning.value = false
  tightenOpen.value = false
  message.success('共享权限收紧任务已完成')
}

function exportLogs() {
  if (!logs.value.length) return message.info('暂无可导出的操作记录')
  const header = ['时间', '操作', '状态', '表格', '文件ID', '目标账号', '详情']
  const rows = logs.value.map((log) => [
    dayjs(log.time).format('YYYY-MM-DD HH:mm:ss'),
    log.action,
    log.status,
    log.fileName,
    log.fileId,
    log.target || '',
    log.detail,
  ])
  const csv = [header, ...rows]
    .map((row) => row.map(escapeCsvCell).join(','))
    .join('\n')
  const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `sheet-guardian-${dayjs().format('YYYYMMDD-HHmmss')}.csv`
  anchor.click()
  URL.revokeObjectURL(url)
}

function clearLogs() {
  Modal.confirm({
    title: '清空操作记录？',
    content: '该操作只清除本机扩展内的日志，不会影响任何 Google 表格权限。',
    okText: '清空',
    okType: 'danger',
    cancelText: '取消',
    onOk: () => {
      logs.value = []
    },
  })
}

function showHelp() {
  Modal.info({
    title: '设置与帮助',
    width: 520,
    okText: '知道了',
    content: () =>
      '插件使用当前 Chrome 配置文件中的主 Google 身份。若要切换账号，请使用另一个 Chrome 配置文件。首次安装需要在 Google Cloud 启用 Drive API，并在 manifest.json 中配置 Chrome 扩展 OAuth 客户端 ID。详细步骤请查看项目 README.md。',
  })
}

async function copyFileId(id: string) {
  try {
    await navigator.clipboard.writeText(id)
    message.success('文件 ID 已复制')
  } catch {
    message.error('复制失败，请手动复制文件 ID')
  }
}

function roleLabel(value: string) {
  return { reader: '可查看', commenter: '可评论', writer: '可编辑', owner: '转移所有者' }[value] || value
}

function statusText(status: OperationLog['status']) {
  return { success: '成功', failed: '失败', skipped: '已跳过' }[status]
}

function isOperationLog(value: unknown): value is OperationLog {
  if (!value || typeof value !== 'object') return false
  const log = value as Record<string, unknown>
  const actions: OperationLog['action'][] = ['授权', 'Sheet 保护权限', '转移所有权', '开启共享', '收紧共享', '移至回收站']
  const statuses: OperationLog['status'][] = ['success', 'failed', 'skipped']
  return (
    typeof log.id === 'string' &&
    typeof log.time === 'string' &&
    typeof log.action === 'string' &&
    actions.includes(log.action as OperationLog['action']) &&
    typeof log.status === 'string' &&
    statuses.includes(log.status as OperationLog['status']) &&
    typeof log.fileId === 'string' &&
    typeof log.fileName === 'string' &&
    (log.target === undefined || typeof log.target === 'string') &&
    typeof log.detail === 'string'
  )
}

onMounted(async () => {
  try {
    const stored = localStorage.getItem('sheet-guardian-logs')
    if (stored) {
      const parsed: unknown = JSON.parse(stored)
      if (Array.isArray(parsed)) logs.value = parsed.filter(isOperationLog).slice(0, 500)
    }
  } catch {
    // Ignore malformed local-only history.
  }

  if (!demoMode) {
    try {
      profile.value = await connectGoogle(false)
      connected.value = true
      await scanSheets()
    } catch {
      connected.value = false
    }
  }
})
</script>

<template>
  <a-config-provider
    :locale="zhCN"
    :theme="{
      token: {
        colorPrimary: '#2764ff',
        colorInfo: '#2764ff',
        borderRadius: 10,
        fontFamily: `'Inter', 'SF Pro Display', 'PingFang SC', 'Microsoft YaHei', sans-serif`,
      },
    }"
  >
    <div class="app-shell">
      <aside class="sidebar">
        <div class="brand">
          <div class="brand-mark"><img src="/icons/icon-128.png" alt="" /></div>
          <div>
            <div class="brand-name">Sheet Guardian</div>
            <div class="brand-subtitle">Google 表格权限管家</div>
          </div>
        </div>

        <nav class="side-nav">
          <button :class="['nav-item', { active: activeTab === 'grant' }]" @click="activeTab = 'grant'">
            <SendOutlined /><span>批量授权</span>
          </button>
          <button :class="['nav-item', { active: activeTab === 'audit' }]" @click="activeTab = 'audit'">
            <EyeOutlined /><span>批量查权限</span>
            <span v-if="auditResults.length" class="nav-count">{{ auditResults.length }}</span>
          </button>
          <button :class="['nav-item', { active: activeTab === 'scan' }]" @click="activeTab = 'scan'">
            <FileSearchOutlined /><span>网盘扫描</span>
            <span v-if="sheets.length" class="nav-count">{{ sheets.length }}</span>
          </button>
          <button :class="['nav-item', { active: activeTab === 'history' }]" @click="activeTab = 'history'">
            <HistoryOutlined /><span>操作记录</span>
            <span v-if="logs.length" class="nav-count">{{ logs.length }}</span>
          </button>
        </nav>

        <div class="sidebar-spacer" />
        <div class="safety-note">
          <SafetyOutlined />
          <div>
            <strong>安全默认值</strong>
            <p>收紧共享时不会移除已明确添加的用户或群组。</p>
          </div>
        </div>
        <button class="nav-item muted" @click="showHelp"><SettingOutlined /><span>设置与帮助</span></button>
      </aside>

      <main class="main-area">
        <header class="topbar">
          <div>
            <h1 v-if="activeTab === 'grant'">批量授权</h1>
            <h1 v-else-if="activeTab === 'audit'">批量查看表格权限</h1>
            <h1 v-else-if="activeTab === 'scan'">Google Drive 文件扫描与共享治理</h1>
            <h1 v-else>操作记录</h1>
            <p v-if="activeTab === 'grant'">支持 Google Drive 文件组合授权或按顺序一一对应账号</p>
            <p v-else-if="activeTab === 'audit'">粘贴 Drive 文件链接或 ID，集中查看所有者与共享账号</p>
            <p v-else-if="activeTab === 'scan'">扫描当前账号可访问的全部文件，并按类型、所有者和共享范围筛选</p>
            <p v-else>查看每一项权限变更的执行结果</p>
          </div>

          <div class="account-area">
            <a-tag v-if="demoMode" color="gold" class="demo-tag">网页演示模式</a-tag>
            <template v-if="connected">
              <div class="account-pill">
                <div class="google-avatar"><GoogleOutlined /></div>
                <div class="account-copy">
                  <span>当前 Google 账号</span>
                  <strong>{{ profile.email || 'Chrome 当前账号' }}</strong>
                </div>
                <a-button type="text" size="small" title="退出授权" @click="handleDisconnect"><LogoutOutlined /></a-button>
              </div>
            </template>
            <a-button v-else type="primary" size="large" :loading="connecting" @click="handleConnect">
              <GoogleOutlined />连接 Google 账号
            </a-button>
          </div>
        </header>

        <div v-if="demoMode" class="demo-banner">
          <ExclamationCircleFilled />
          <span>当前用示例数据预览界面。将 <code>dist/</code> 加载为 Chrome 扩展后，会使用浏览器中的 Google 身份。</span>
        </div>

        <section v-if="activeTab === 'grant'" class="page-content grant-layout">
          <div class="grant-mode-bar">
            <div>
              <strong>授权方式</strong>
              <span v-if="grantMappingMode === 'matrix'">每个文件都授权给右侧全部账号</span>
              <span v-else-if="grantMappingMode === 'paired'">严格按输入顺序配对：第 1 个文件 → 第 1 个账号，依次对应</span>
              <span v-else>不填写邮箱，直接为每个文件开启“知道链接的任何人”权限</span>
            </div>
            <a-radio-group v-model:value="grantMappingMode" button-style="solid" size="large">
              <a-radio-button value="matrix">组合授权</a-radio-button>
              <a-radio-button value="paired">一一对应</a-radio-button>
              <a-radio-button value="public">公开链接</a-radio-button>
            </a-radio-group>
            <a-tag
              v-if="grantMappingMode === 'paired' && (parsedFileIds.length || parsedEmails.length)"
              :color="pairedCountMatches ? 'green' : 'red'"
            >
              {{ parsedFileIds.length }} 个文件 / {{ parsedEmails.length }} 个账号
              · {{ pairedCountMatches ? '数量一致' : '数量必须相同' }}
            </a-tag>
          </div>

          <div class="workflow-card">
            <div class="section-heading">
              <div class="step-badge">1</div>
              <div>
                <h2>{{ pairedMode ? '按顺序添加目标文件' : '添加目标文件' }}</h2>
                <p>{{ pairedMode ? '每行一个文件；顺序将与右侧账号逐项对应' : '支持表格、文档、文件夹等 Drive 链接或文件 ID' }}</p>
              </div>
            </div>
            <a-textarea
              v-model:value="linkInput"
              class="input-area"
              :rows="7"
              placeholder="https://docs.google.com/spreadsheets/d/1AbC.../edit&#10;https://docs.google.com/spreadsheets/d/1Def.../edit"
            />
            <div class="input-meta">
              <span><LinkOutlined /> 已识别 {{ parsedFileIds.length }} 个{{ pairedMode ? '有序' : '唯一' }}文件</span>
              <span v-if="unrecognizedLinkCount" class="warning-text">{{ unrecognizedLinkCount }} 项无法识别</span>
              <span class="meta-hint">{{ pairedMode ? '保留顺序及重复项' : '自动去重' }}</span>
            </div>
            <a-button
              size="large"
              :loading="resolving"
              :disabled="!parsedFileIds.length"
              @click="resolveLinks"
            >
              <AuditOutlined />解析并检查权限
            </a-button>

            <div v-if="resolvedTargets.length" class="resolved-list">
              <div v-for="(target, index) in resolvedTargets" :key="target.id" :class="['resolved-item', { error: target.error }]">
                <CheckCircleFilled v-if="target.file" />
                <ExclamationCircleFilled v-else />
                <div>
                  <strong>{{ pairedMode ? `${index + 1}. ` : '' }}{{ target.file?.name || '无法读取' }}</strong>
                  <span>{{ target.error || target.id }}</span>
                </div>
                <a-tag v-if="ownershipMode && target.file" :color="canTransferOwnership(target.file) ? 'green' : 'orange'">
                  {{ canTransferOwnership(target.file) ? '可转移' : target.file.driveId ? '共享云端硬盘' : '非当前账号所有' }}
                </a-tag>
                <a-tag v-else-if="target.file?.capabilities?.canShare === false" color="orange">无共享权限</a-tag>
                <a-tag v-else-if="target.file" color="green">可授权</a-tag>
              </div>
            </div>
          </div>

          <div class="workflow-card">
            <div class="section-heading">
              <div class="step-badge">2</div>
              <div>
                <h2>{{ publicGrantMode ? '设置公开链接权限' : pairedMode ? '按顺序添加账号' : ownershipMode ? '设置新所有者' : '添加账号' }}</h2>
                <p v-if="publicGrantMode">无需邮箱，持有文件链接的任何人都能获得所选权限</p>
                <p v-else-if="pairedMode">每行一个邮箱；第 N 个账号只应用到左侧第 N 个文件</p>
                <p v-else>{{ ownershipMode ? '一次可将多个文件转移给同一个目标账号' : '输入需要获得权限的 Google 账号邮箱' }}</p>
              </div>
            </div>
            <a-textarea
              v-if="!publicGrantMode"
              v-model:value="emailInput"
              class="input-area"
              :rows="7"
              placeholder="alice@example.com&#10;bob@example.com&#10;team@example.com"
            />
            <div v-if="!publicGrantMode" class="input-meta">
              <span><MailOutlined /> 已识别 {{ parsedEmails.length }} 个{{ pairedMode ? '有序' : '唯一' }}邮箱</span>
              <span v-if="invalidEmailCount" class="warning-text">{{ invalidEmailCount }} 项格式有误</span>
              <span class="meta-hint">{{ pairedMode ? '按顺序配对，不会交叉授权' : ownershipMode ? '组合转移仅允许 1 个邮箱' : '不区分大小写' }}</span>
            </div>

            <div v-if="!publicGrantMode" class="setting-grid">
              <label>
                <span>权限角色</span>
                <a-select v-model:value="role" size="large">
                  <a-select-option value="reader">可查看</a-select-option>
                  <a-select-option value="commenter">可评论</a-select-option>
                  <a-select-option value="writer">可编辑</a-select-option>
                  <a-select-option value="owner">转移所有者（高风险）</a-select-option>
                </a-select>
              </label>
              <label class="switch-setting">
                <span>{{ ownershipMode ? '必须发送转移通知' : '发送 Google 通知邮件' }}</span>
                <a-switch v-model:checked="notifyEmail" :disabled="ownershipMode" />
              </label>
            </div>
            <a-alert
              v-if="!publicGrantMode && ownershipMode"
              class="ownership-alert"
              type="error"
              show-icon
              message="所有权转移后，当前账号将不再是这些文件的所有者"
              description="只支持当前账号拥有的“我的云端硬盘”文件。共享云端硬盘文件无法转移；个人 Gmail 需要目标账号接受邀请。"
            />
            <a-input
              v-if="!publicGrantMode && notifyEmail"
              v-model:value="emailMessage"
              size="large"
              :maxlength="500"
              placeholder="可选：在通知邮件中附上一句话"
            />

            <div v-if="publicGrantMode" class="public-grant-panel">
              <div class="public-grant-icon"><LinkOutlined /></div>
              <div class="public-grant-copy">
                <strong>无需添加邮箱</strong>
                <span>Google Drive 会创建或更新 anyone 链接权限</span>
              </div>
              <label class="grant-link-role-setting">
                <span>公开链接角色</span>
                <a-select v-model:value="grantLinkSharingRole" size="large">
                  <a-select-option value="reader">知道链接的任何人 · 可查看</a-select-option>
                  <a-select-option value="commenter">知道链接的任何人 · 可评论</a-select-option>
                  <a-select-option value="writer">知道链接的任何人 · 可编辑（高风险）</a-select-option>
                </a-select>
              </label>
              <a-alert
                :type="grantLinkSharingRole === 'writer' ? 'error' : 'warning'"
                show-icon
                :message="grantLinkSharingRole === 'writer' ? '任何拿到链接的人都可以修改内容' : `任何拿到链接的人都可以${roleLabel(grantLinkSharingRole)}`"
                description="不需要填写邮箱或发送通知。Workspace 管理策略可能禁止公开共享；已有公开链接角色会直接更新。"
              />
            </div>

            <div v-if="!publicGrantMode && !ownershipMode" :class="['protected-sheet-option', { active: grantEnableLinkSharing }]">
              <a-checkbox v-model:checked="grantEnableLinkSharing">
                <span class="protected-sheet-option-copy">
                  <strong>添加邮箱时同时开启链接共享（可选）</strong>
                  <small>默认关闭；为每个目标文件开启“知道链接的任何人”</small>
                </span>
              </a-checkbox>
              <div v-if="grantEnableLinkSharing" class="protected-sheet-fields">
                <label class="grant-link-role-setting">
                  <span>链接访问角色</span>
                  <a-select v-model:value="grantLinkSharingRole" size="large">
                    <a-select-option value="reader">知道链接的任何人 · 可查看</a-select-option>
                    <a-select-option value="commenter">知道链接的任何人 · 可评论</a-select-option>
                    <a-select-option value="writer">知道链接的任何人 · 可编辑（高风险）</a-select-option>
                  </a-select>
                </label>
                <a-alert
                  :type="grantLinkSharingRole === 'writer' ? 'error' : 'warning'"
                  show-icon
                  :message="`将开启公开链接${roleLabel(grantLinkSharingRole)}权限`"
                  description="持有链接的人将不需要单独添加邮箱；此操作不会代替 IMPORTRANGE 首次连接时的“允许访问”确认。"
                />
              </div>
            </div>

            <div v-if="!publicGrantMode && !ownershipMode" :class="['protected-sheet-option', { active: syncProtectedSheetEditors }]">
              <a-checkbox v-model:checked="syncProtectedSheetEditors">
                <span class="protected-sheet-option-copy">
                  <strong>同时添加到 Sheet 保护权限（可选）</strong>
                  <small>默认关闭；只对 Google 表格执行，其他文件自动跳过</small>
                </span>
              </a-checkbox>
              <div v-if="syncProtectedSheetEditors" class="protected-sheet-fields">
                <a-input
                  v-model:value="sheetKeywordInput"
                  size="large"
                  placeholder="Sheet 关键词，例如：财务、审批、工资"
                >
                  <template #prefix><SearchOutlined /></template>
                </a-input>
                <div class="sheet-keyword-meta">
                  <span>已识别 {{ sheetKeywords.length }} 个关键词</span>
                  <small>逗号、空格或换行分隔；标题包含任一关键词即匹配</small>
                </div>
                <a-alert
                  type="info"
                  show-icon
                  message="只追加已有保护规则的编辑者"
                  description="不会新建保护规则，也不会覆盖原有用户、群组或域设置；仅警告规则与当前账号无权修改的规则会跳过。"
                />
              </div>
            </div>
          </div>

          <aside class="execution-card">
            <div :class="['execution-icon', { danger: ownershipMode || (publicGrantMode && grantLinkSharingRole === 'writer') }]"><CloudSyncOutlined /></div>
            <h2>{{ publicGrantMode ? '公开权限预览' : ownershipMode ? '所有权转移预览' : '执行预览' }}</h2>
            <p>{{ publicGrantMode ? '不需要邮箱，将逐个修改文件的链接共享' : pairedMode ? '系统只执行同序号的文件与账号组合' : ownershipMode ? '确认文件与唯一目标账号后再继续' : '检查无误后再开始批量授权' }}</p>

            <div class="preview-metric">
              <span>{{ ownershipMode ? '可转移文件' : '有效文件' }}</span><strong>{{ ownershipMode ? transferableTargets.length : validResolvedTargets.length }}</strong>
            </div>
            <div class="preview-metric">
              <span>{{ publicGrantMode ? '访问对象' : '目标账号' }}</span>
              <strong :class="{ small: publicGrantMode }">{{ publicGrantMode ? '任何持有链接的人' : parsedEmails.length }}</strong>
            </div>
            <div class="preview-metric">
              <span>权限角色</span><strong class="small">{{ roleLabel(publicGrantMode ? grantLinkSharingRole : role) }}</strong>
            </div>
            <div class="preview-total">
              <span>{{ publicGrantMode ? '预计公开' : ownershipMode ? '预计转移' : '预计处理' }}</span>
              <strong>{{ ownershipMode ? permissionMatrixSize : executionTaskCount }}</strong>
              <small>{{ ownershipMode ? '个文件' : executionTaskUnit }}</small>
            </div>

            <a-alert
              v-if="publicGrantMode && validResolvedTargets.length"
              :type="grantLinkSharingRole === 'writer' ? 'error' : 'warning'"
              show-icon
              :message="grantLinkSharingRole === 'writer' ? '公开编辑权限风险很高，请确认文件内容可被任何人修改' : '此操作会扩大文件访问范围'"
            />
            <a-alert
              v-else-if="validResolvedTargets.length && parsedEmails.length && !ownershipMode"
              :type="pairedMode && !pairedCountMatches ? 'error' : 'info'"
              show-icon
              :message="pairedMode ? (pairedCountMatches ? '将按序号一一授权，不会交叉组合' : '文件数与账号数必须完全相同') : '同一账号已有相同权限时会自动跳过'"
            />
            <a-alert
              v-else-if="ownershipMode && validResolvedTargets.length"
              type="warning"
              show-icon
              :message="`${transferableTargets.length} 份可转移，${validResolvedTargets.length - transferableTargets.length} 份将跳过`"
            />
            <a-button
              type="primary"
              :danger="ownershipMode || (publicGrantMode && grantLinkSharingRole === 'writer')"
              size="large"
              block
              :loading="granting"
              :disabled="
                !validResolvedTargets.length ||
                (!publicGrantMode && !parsedEmails.length) ||
                (!publicGrantMode && pairedMode && !pairedCountMatches) ||
                (ownershipMode && ((!pairedMode && parsedEmails.length !== 1) || !transferableTargets.length))
              "
              @click="requestBatchExecution"
            >
              <SendOutlined />{{ publicGrantMode ? `批量开启公开${roleLabel(grantLinkSharingRole)}` : ownershipMode ? '审核并转移所有权' : '开始批量授权' }}
            </a-button>
          </aside>
        </section>

        <section v-else-if="activeTab === 'audit'" class="page-content audit-page">
          <div class="audit-input-card">
            <div class="audit-input-copy">
              <div class="section-heading">
                <div class="step-badge"><EyeOutlined /></div>
                <div>
                  <h2>粘贴需要查询的文件</h2>
                  <p>支持表格、文档、幻灯片、文件夹等 Drive 链接或文件 ID</p>
                </div>
              </div>
              <a-textarea
                v-model:value="auditInput"
                class="input-area audit-textarea"
                :rows="5"
                placeholder="https://docs.google.com/spreadsheets/d/1AbC.../edit&#10;https://docs.google.com/spreadsheets/d/1Def.../edit&#10;1GhI..."
              />
              <div class="input-meta">
                <span><LinkOutlined /> 已识别 {{ parsedAuditFileIds.length }} 个唯一文件</span>
                <span v-if="unrecognizedAuditLinkCount" class="warning-text">
                  {{ unrecognizedAuditLinkCount }} 项无法识别
                </span>
                <span class="meta-hint">只读查询，不会修改权限</span>
              </div>
            </div>
            <div class="audit-actions">
              <div class="audit-shield"><SafetyOutlined /></div>
              <strong>安全读取</strong>
              <p>使用当前 Chrome Google 身份读取 Drive 返回的所有者和权限信息。</p>
              <a-button
                type="primary"
                size="large"
                block
                :loading="auditLoading"
                :disabled="!parsedAuditFileIds.length"
                @click="inspectLinks"
              >
                <SearchOutlined />批量查询权限
              </a-button>
              <div class="audit-secondary-actions">
                <a-button :disabled="!auditResults.length" @click="copyAuditResults"><CopyOutlined />复制结果</a-button>
                <a-button :disabled="!auditInput && !auditResults.length" @click="clearAudit">清空</a-button>
              </div>
            </div>
          </div>

          <div v-if="auditResults.length" class="audit-summary">
            <div><span>查询总数</span><strong>{{ auditResults.length }}</strong></div>
            <div><span>读取成功</span><strong class="success">{{ validAuditResults.length }}</strong></div>
            <div><span>公开链接</span><strong class="danger">{{ auditPublicCount }}</strong></div>
            <div><span>受限表格</span><strong>{{ auditRestrictedCount }}</strong></div>
            <div><span>无法访问</span><strong>{{ auditResults.length - validAuditResults.length }}</strong></div>
          </div>

          <section v-if="validAuditResults.length" class="audit-email-panel">
            <header class="audit-email-panel-header">
              <div>
                <h2>按邮箱复制文件链接</h2>
                <p>当前登录邮箱复制全部读取成功的链接；其他邮箱按 Drive 返回的明确用户或群组权限归集。</p>
              </div>
              <a-input
                v-model:value="auditEmailSearch"
                allow-clear
                placeholder="搜索邮箱或名称"
              >
                <template #prefix><SearchOutlined /></template>
              </a-input>
            </header>

            <div class="audit-email-grid">
              <article v-if="showCurrentAuditGroup" class="audit-email-card current-account">
                <div class="audit-email-avatar"><GoogleOutlined /></div>
                <div class="audit-email-copy">
                  <strong>{{ profile.email || '当前登录邮箱' }}</strong>
                  <small>当前身份成功读取的全部表格</small>
                  <div class="audit-email-tags">
                    <a-tag color="blue">当前账号</a-tag>
                    <a-tag>有效访问</a-tag>
                  </div>
                </div>
                <div class="audit-email-action">
                  <strong>{{ currentAccountAuditFiles.length }}</strong>
                  <span>个链接</span>
                  <a-button
                    type="primary"
                    size="small"
                    @click="copyAuditEmailLinks(profile.email || '当前登录邮箱', currentAccountAuditFiles)"
                  >
                    <CopyOutlined />复制全部
                  </a-button>
                </div>
              </article>

              <article v-for="group in filteredAuditEmailGroups" :key="group.email" class="audit-email-card">
                <div class="audit-email-avatar"><TeamOutlined v-if="group.types.includes('group')" /><MailOutlined v-else /></div>
                <div class="audit-email-copy">
                  <strong>{{ group.email }}</strong>
                  <small>{{ group.displayName || (group.types.includes('group') ? 'Google 群组' : 'Google 账号') }}</small>
                  <div class="audit-email-tags">
                    <a-tag v-if="group.types.includes('group')" color="purple">群组</a-tag>
                    <a-tag v-else color="cyan">账号</a-tag>
                    <a-tag v-for="groupRole in group.roles" :key="groupRole">
                      {{ permissionRoleLabel(groupRole) }}
                    </a-tag>
                  </div>
                </div>
                <div class="audit-email-action">
                  <strong>{{ group.files.length }}</strong>
                  <span>个链接</span>
                  <a-button size="small" @click="copyAuditEmailLinks(group.email, group.files)">
                    <CopyOutlined />复制全部
                  </a-button>
                </div>
              </article>

              <div v-if="!showCurrentAuditGroup && !filteredAuditEmailGroups.length" class="audit-email-empty">
                没有匹配的邮箱
              </div>
            </div>
          </section>

          <div v-if="auditResults.length" class="audit-results">
            <article
              v-for="target in auditResults"
              :key="target.id"
              :class="['audit-result-card', { error: target.error }]"
            >
              <template v-if="target.file">
                <header class="audit-result-header">
                  <div class="sheet-icon">S</div>
                  <div class="audit-result-title">
                    <a :href="spreadsheetLink(target.file)" target="_blank" rel="noopener noreferrer">{{ target.file.name }}</a>
                    <button type="button" @click="copyFileId(target.file.id)">
                      {{ target.file.id }} <CopyOutlined />
                    </button>
                  </div>
                  <div class="tag-row audit-status-tags">
                    <a-tag v-if="sharingSummary(target.file).anyone" color="red">公开链接</a-tag>
                    <a-tag v-if="sharingSummary(target.file).domain" color="gold">域内共享</a-tag>
                    <a-tag v-if="!sharingSummary(target.file).risky" color="green">受限</a-tag>
                    <a-tag :color="target.file.capabilities?.canShare === false ? 'default' : 'blue'">
                      {{ target.file.capabilities?.canShare === false ? '当前账号不可管理' : '当前账号可管理' }}
                    </a-tag>
                  </div>
                </header>

                <div class="audit-facts">
                  <div>
                    <span>所有者</span>
                    <strong>{{ target.file.driveId ? '共享云端硬盘' : target.file.owners?.[0]?.displayName || '所有者账号' }}</strong>
                    <small>{{ ownerText(target.file) }}</small>
                  </div>
                  <div>
                    <span>指定账号 / 群组</span>
                    <strong>{{ namedPermissions(target.file).length }}</strong>
                    <small>包括所有者和直接共享账号</small>
                  </div>
                  <div>
                    <span>最近修改</span>
                    <strong>{{ target.file.modifiedTime ? dayjs(target.file.modifiedTime).format('YYYY-MM-DD') : '—' }}</strong>
                    <small>{{ target.file.modifiedTime ? dayjs(target.file.modifiedTime).format('HH:mm') : '未返回时间' }}</small>
                  </div>
                </div>

                <div class="audit-permission-section">
                  <div class="audit-permission-heading">
                    <strong>权限明细</strong>
                    <span>共 {{ target.file.permissions?.length || 0 }} 项</span>
                  </div>
                  <div v-if="target.file.permissions?.length" class="audit-permission-list">
                    <div v-for="permission in target.file.permissions" :key="permission.id" class="audit-permission-item">
                      <span :class="['permission-type', permission.type]">{{ permissionTypeLabel(permission.type) }}</span>
                      <div>
                        <strong>{{ permissionTarget(permission) }}</strong>
                        <small v-if="permission.displayName && permission.emailAddress">{{ permission.displayName }}</small>
                      </div>
                      <a-tag :color="permission.role === 'owner' ? 'purple' : permission.type === 'anyone' ? 'red' : 'blue'">
                        {{ permissionRoleLabel(permission.role) }}
                      </a-tag>
                      <a-tag v-if="isInheritedPermission(permission)">继承</a-tag>
                      <a-tag v-if="permission.pendingOwner" color="orange">待接受所有权</a-tag>
                    </div>
                  </div>
                  <div v-else class="audit-no-permission">Google Drive 未向当前账号返回权限明细</div>
                </div>
              </template>
              <template v-else>
                <div class="audit-error-icon"><ExclamationCircleFilled /></div>
                <div>
                  <strong>无法读取此文件</strong>
                  <p>{{ target.error }}</p>
                  <small>{{ target.id }}</small>
                </div>
              </template>
            </article>
          </div>

          <div v-else-if="!auditLoading" class="audit-empty">
            <EyeOutlined />
            <strong>粘贴多条 Drive 文件链接后即可集中查看</strong>
            <span>查询结果会显示所有者、公开范围、共享账号和对应角色</span>
          </div>
        </section>

        <section v-else-if="activeTab === 'scan'" class="page-content scan-page">
          <div class="stats-grid">
            <div class="stat-card">
              <div class="stat-icon blue"><AppstoreOutlined /></div>
              <div><span>已扫描文件</span><strong>{{ sheets.length }}</strong></div>
            </div>
            <div class="stat-card">
              <div class="stat-icon green"><FileSearchOutlined /></div>
              <div><span>Google 表格</span><strong>{{ spreadsheetFileCount }}</strong></div>
            </div>
            <div class="stat-card">
              <div class="stat-icon red"><LinkOutlined /></div>
              <div><span>公开链接</span><strong>{{ publicSheetCount }}</strong></div>
            </div>
            <div class="stat-card">
              <div class="stat-icon gold"><TeamOutlined /></div>
              <div><span>域内共享</span><strong>{{ domainSheetCount }}</strong></div>
            </div>
            <div class="stat-card">
              <div class="stat-icon green"><FileProtectOutlined /></div>
              <div><span>仅指定账号</span><strong>{{ safeSheetCount }}</strong></div>
            </div>
            <div class="stat-card">
              <div class="stat-icon purple"><GoogleOutlined /></div>
              <div><span>当前账号拥有</span><strong>{{ ownedSheetCount }}</strong></div>
            </div>
          </div>

          <div class="table-card">
            <div class="table-toolbar">
              <div>
                <h2>Google Drive 文件清单</h2>
                <p>包含表格、文档、幻灯片、文件夹、PDF、图片及其他文件；可批量授权、治理共享或移入回收站</p>
              </div>
              <div class="toolbar-actions">
                <a-input v-model:value="searchText" class="toolbar-search" size="large" placeholder="搜索文件、所有者姓名或邮箱" allow-clear>
                  <template #prefix><SearchOutlined /></template>
                </a-input>
                <a-button size="large" :loading="scanLoading" @click="scanSheets"><ReloadOutlined />重新扫描</a-button>
                <a-button
                  v-if="selectedRowKeys.length"
                  type="primary"
                  size="large"
                  @click="sendSelectedToGrant"
                >
                  <MailOutlined />批量添加邮箱 ({{ selectedRowKeys.length }})
                </a-button>
                <a-button
                  v-if="selectedRowKeys.length"
                  class="open-sharing-button"
                  size="large"
                  @click="openEnableSharingModal"
                >
                  <LinkOutlined />开启共享 ({{ selectedRowKeys.length }})
                </a-button>
                <a-button
                  v-if="selectedRowKeys.length"
                  class="clear-selection-button"
                  size="large"
                  @click="clearSelectedSheets"
                >
                  <StopOutlined />取消全部选择
                </a-button>
                <a-button size="large" type="primary" danger :disabled="!selectedRowKeys.length" @click="openTightenModal">
                  <LockOutlined />收紧共享<span v-if="selectedRowKeys.length"> ({{ selectedRowKeys.length }})</span>
                </a-button>
                <a-button size="large" danger :disabled="!selectedRowKeys.length" @click="openTrashModal">
                  <DeleteOutlined />删除文件<span v-if="selectedRowKeys.length"> ({{ selectedRowKeys.length }})</span>
                </a-button>
              </div>
            </div>

            <div class="scan-filter-bar">
              <div class="scan-filter-control">
                <span><FileSearchOutlined />文件类型</span>
                <a-select v-model:value="fileTypeFilter" size="middle" popup-class-name="scan-filter-dropdown">
                  <a-select-option value="all">全部文件</a-select-option>
                  <a-select-option value="spreadsheet">Google 表格</a-select-option>
                  <a-select-option value="document">Google 文档</a-select-option>
                  <a-select-option value="presentation">Google 幻灯片</a-select-option>
                  <a-select-option value="folder">文件夹</a-select-option>
                  <a-select-option value="pdf">PDF</a-select-option>
                  <a-select-option value="image">图片</a-select-option>
                  <a-select-option value="other">其他文件</a-select-option>
                </a-select>
              </div>
              <div class="scan-filter-control">
                <span><FilterOutlined />权限范围</span>
                <a-select v-model:value="sharingFilter" size="middle" popup-class-name="scan-filter-dropdown">
                  <a-select-option value="all">全部权限</a-select-option>
                  <a-select-option value="public">公开链接</a-select-option>
                  <a-select-option value="domain">域内共享</a-select-option>
                  <a-select-option value="restricted">受限</a-select-option>
                </a-select>
              </div>
              <div class="scan-filter-control">
                <span><GoogleOutlined />所有权</span>
                <a-select v-model:value="ownershipFilter" size="middle" popup-class-name="scan-filter-dropdown">
                  <a-select-option value="all">全部文件</a-select-option>
                  <a-select-option value="mine">当前账号拥有</a-select-option>
                  <a-select-option value="others">其他所有者</a-select-option>
                </a-select>
              </div>
              <div class="scan-filter-control">
                <span><SortAscendingOutlined />修改日期</span>
                <a-select v-model:value="modifiedSort" size="middle" popup-class-name="scan-filter-dropdown">
                  <a-select-option value="newest">最新优先</a-select-option>
                  <a-select-option value="oldest">最旧优先</a-select-option>
                </a-select>
              </div>
              <button type="button" class="select-owned-button" :disabled="!ownedSheetCount" @click="selectAllOwnedSheets">
                选择我拥有的全部 ({{ ownedSheetCount }})
              </button>
              <span class="scan-filter-result">显示 {{ filteredSheets.length }} / {{ sheets.length }} 个文件</span>
              <button v-if="hasActiveScanFilters" type="button" class="reset-filter-button" @click="resetScanFilters">
                重置筛选
              </button>
            </div>

            <div v-if="scanLoading" class="scan-progress">
              <a-spin size="small" /> 正在读取 Google Drive{{ scanCount ? `，已发现 ${scanCount} 个文件` : '' }}…
            </div>

            <a-table
              row-key="id"
              :columns="columns"
              :data-source="filteredSheets"
              :loading="scanLoading"
              :row-selection="{
                selectedRowKeys,
                preserveSelectedRowKeys: true,
                columnWidth: 54,
                onChange: (keys: (string | number)[]) => (selectedRowKeys = keys.map(String)),
                getCheckboxProps: (record: SheetFile) => ({
                  disabled: record.capabilities?.canShare === false && record.capabilities?.canTrash === false,
                }),
              }"
              :pagination="tablePagination"
              @change="handleTableChange"
            >
              <template #bodyCell="{ column, record }">
                <template v-if="column.key === 'name'">
                  <div class="file-cell">
                    <div :class="['drive-file-icon', driveFileCategory(record)]">{{ driveFileIconText(record) }}</div>
                    <div>
                      <a :href="spreadsheetLink(record)" target="_blank" rel="noopener noreferrer">{{ record.name }}</a>
                      <a-tag class="file-type-tag">{{ driveFileTypeLabel(record) }}</a-tag>
                      <button title="复制文件 ID" @click="copyFileId(record.id)">
                        {{ record.id.slice(0, 12) }}… <CopyOutlined />
                      </button>
                    </div>
                  </div>
                </template>
                <template v-else-if="column.key === 'owner'">
                  <div class="owner-cell">
                    <strong>{{ record.driveId ? '共享云端硬盘' : record.owners?.[0]?.displayName || '未知' }}</strong>
                    <span>{{ record.driveId ? '归组织所有' : record.owners?.[0]?.emailAddress || (record.ownedByMe ? profile.email : '—') }}</span>
                    <a-tag v-if="isOwnedByCurrentUser(record)" color="blue">当前账号拥有</a-tag>
                  </div>
                </template>
                <template v-else-if="column.key === 'sharing'">
                  <div class="tag-row">
                    <a-tag v-if="sharingSummary(record).anyone" color="red">公开链接 · {{ sharingSummary(record).anyone }}</a-tag>
                    <a-tag v-if="sharingSummary(record).domain" color="gold">域内共享 · {{ sharingSummary(record).domain }}</a-tag>
                    <a-tag v-if="!sharingSummary(record).risky" color="green">受限</a-tag>
                    <a-tag>账号 {{ sharingSummary(record).named }}</a-tag>
                  </div>
                </template>
                <template v-else-if="column.key === 'modified'">
                  <span class="date-cell">{{ record.modifiedTime ? dayjs(record.modifiedTime).format('YYYY-MM-DD HH:mm') : '—' }}</span>
                </template>
              </template>
              <template #emptyText>
                <div class="empty-state">
                  <FileSearchOutlined />
                  <strong>{{ sheets.length ? '没有符合筛选条件的文件' : '还没有扫描结果' }}</strong>
                  <span>{{ sheets.length ? '请调整搜索词、权限范围或日期排序' : '连接 Google 账号后点击“重新扫描”' }}</span>
                </div>
              </template>
            </a-table>
          </div>
        </section>

        <section v-else class="page-content history-page">
          <div class="table-card">
            <div class="table-toolbar">
              <div>
                <h2>最近操作</h2>
                <p>记录保存在本机扩展中，最多保留 500 条</p>
              </div>
              <div class="toolbar-actions">
                <a-button @click="exportLogs"><DownloadOutlined />导出 CSV</a-button>
                <a-button danger :disabled="!logs.length" @click="clearLogs"><DeleteOutlined />清空记录</a-button>
              </div>
            </div>

            <div v-if="logs.length" class="history-list">
              <div v-for="log in logs" :key="log.id" class="history-item">
                <div :class="['history-status', log.status]">
                  <CheckCircleFilled v-if="log.status === 'success'" />
                  <ExclamationCircleFilled v-else-if="log.status === 'failed'" />
                  <StopOutlined v-else />
                </div>
                <div class="history-main">
                  <div class="history-title">
                    <strong>{{ log.action }} · {{ log.fileName }}</strong>
                    <a-tag :color="log.status === 'success' ? 'green' : log.status === 'failed' ? 'red' : 'default'">
                      {{ statusText(log.status) }}
                    </a-tag>
                  </div>
                  <p>{{ log.detail }}<span v-if="log.target"> · {{ log.target }}</span></p>
                  <small><ClockCircleOutlined /> {{ dayjs(log.time).format('YYYY-MM-DD HH:mm:ss') }} · {{ log.fileId }}</small>
                </div>
              </div>
            </div>
            <div v-else class="empty-state large">
              <HistoryOutlined />
              <strong>暂无操作记录</strong>
              <span>完成批量授权、共享治理或删除后，结果会显示在这里</span>
            </div>
          </div>
        </section>
      </main>

      <a-modal
        v-model:open="ownershipOpen"
        title="确认批量转移所有者"
        ok-text="确认转移"
        cancel-text="取消"
        :ok-button-props="{ danger: true, disabled: !ownershipConfirmed }"
        @ok="confirmOwnershipTransfer"
      >
        <div class="modal-copy ownership-modal">
          <div v-if="!pairedMode" class="ownership-target">
            <span>目标所有者</span>
            <strong>{{ ownershipTargetEmail }}</strong>
          </div>
          <div v-else class="paired-transfer-preview">
            <strong>一一对应转移预览</strong>
            <div v-for="(pair, index) in pairedTasks.slice(0, 5)" :key="pair.target.id">
              <span>{{ index + 1 }}. {{ pair.target.file?.name || pair.target.id }}</span>
              <b>→</b>
              <span>{{ pair.email }}</span>
            </div>
            <small v-if="pairedTasks.length > 5">还有 {{ pairedTasks.length - 5 }} 组未显示</small>
          </div>
          <a-alert
            type="error"
            show-icon
            message="这是高风险且通常不可自动撤销的操作"
            :description="
              pairedMode
                ? '每个文件会转移给同序号的目标账号。个人 Gmail 通常需要目标账号接受邀请；Workspace 跨组织转移可能被拒绝。'
                : consumerOwnershipFlow
                ? 'Google 会向目标 Gmail 账号发送所有权转移邀请；目标账号接受前，当前账号仍是所有者。'
                : 'Google Workspace 同组织转移通常会立即生效；当前账号会降为编辑者。跨组织转移可能被 Google 拒绝。'
            "
          />
          <div class="ownership-counts">
            <div><span>已解析</span><strong>{{ validResolvedTargets.length }}</strong></div>
            <div><span>可转移</span><strong>{{ transferableTargets.length }}</strong></div>
            <div><span>将跳过</span><strong>{{ validResolvedTargets.length - transferableTargets.length }}</strong></div>
          </div>
          <a-checkbox v-model:checked="ownershipConfirmed" class="ownership-confirm">
            我已核对{{ pairedMode ? '每组文件与目标邮箱' : '目标邮箱' }}，并理解转移后当前账号可能失去所有者权限
          </a-checkbox>
        </div>
      </a-modal>

      <a-modal
        v-model:open="trashOpen"
        title="批量删除文件"
        :confirm-loading="trashRunning"
        ok-text="移入回收站"
        cancel-text="取消"
        :ok-button-props="{ danger: true, disabled: !trashConfirmed || !selectedTrashableCount }"
        @ok="executeTrash"
      >
        <div class="modal-copy trash-modal">
          <div class="trash-warning-icon"><DeleteOutlined /></div>
          <div class="trash-heading">
            <strong>所选文件将移入 Google Drive 回收站</strong>
            <p>这不是永久删除；在回收站被清空前，可以从 Google Drive 恢复。</p>
          </div>
          <div class="ownership-counts">
            <div><span>已选择</span><strong>{{ selectedSheets.length }}</strong></div>
            <div><span>可移入回收站</span><strong>{{ selectedTrashableCount }}</strong></div>
            <div><span>无权操作</span><strong>{{ selectedSheets.length - selectedTrashableCount }}</strong></div>
          </div>
          <div class="trash-file-preview">
            <div v-for="file in selectedSheets.slice(0, 5)" :key="file.id">
              <span :class="['drive-file-icon', 'compact', driveFileCategory(file)]">{{ driveFileIconText(file) }}</span>
              <strong>{{ file.name }}</strong>
              <small>{{ file.id }}</small>
            </div>
            <p v-if="selectedSheets.length > 5">还有 {{ selectedSheets.length - 5 }} 个文件未显示</p>
          </div>
          <a-alert
            type="error"
            show-icon
            message="移入回收站后，相关链接和 IMPORTRANGE 可能立即停止工作"
            description="共享账号可能无法继续访问这些文件。插件不会清空回收站，也不会执行永久删除。"
          />
          <a-checkbox v-model:checked="trashConfirmed" class="ownership-confirm">
            我已核对所选文件，并确认将它们移入回收站
          </a-checkbox>
        </div>
      </a-modal>

      <a-modal
        v-model:open="openSharingOpen"
        title="批量开启链接共享"
        :confirm-loading="openSharingRunning"
        ok-text="确认开启"
        cancel-text="取消"
        :ok-button-props="{ danger: openSharingRole === 'writer', disabled: !openSharingConfirmed }"
        @ok="executeEnableSharing"
      >
        <div class="modal-copy open-sharing-modal">
          <div class="modal-safety"><LinkOutlined /><strong>已选 {{ selectedSheets.length }} 个文件</strong></div>
          <p>为所选文件开启“知道链接的任何人”。已有相同权限会跳过，已有其他链接角色会更新。</p>
          <label class="open-sharing-role">
            <span>链接访问角色</span>
            <a-select v-model:value="openSharingRole" size="large">
              <a-select-option value="reader">可查看</a-select-option>
              <a-select-option value="commenter">可评论</a-select-option>
              <a-select-option value="writer">可编辑（高风险）</a-select-option>
            </a-select>
          </label>
          <div class="ownership-counts">
            <div><span>已选择</span><strong>{{ selectedSheets.length }}</strong></div>
            <div><span>已有公开链接</span><strong>{{ selectedExistingPublicCount }}</strong></div>
            <div><span>当前受限</span><strong>{{ selectedSheets.length - selectedExistingPublicCount }}</strong></div>
          </div>
          <a-alert
            :type="openSharingRole === 'writer' ? 'error' : 'warning'"
            show-icon
            :message="`持有链接的任何人都将${roleLabel(openSharingRole)}`"
            description="此操作会扩大数据访问范围，但不会自动完成 IMPORTRANGE 首次连接的“允许访问”确认。Workspace 管理策略可能禁止公开共享。"
          />
          <a-checkbox v-model:checked="openSharingConfirmed" class="ownership-confirm">
            我确认所选文件可以开放给持有链接的任何人
          </a-checkbox>
        </div>
      </a-modal>

      <a-modal
        v-model:open="tightenOpen"
        title="批量收紧共享权限"
        :confirm-loading="tightenRunning"
        ok-text="确认收紧"
        cancel-text="取消"
        :ok-button-props="{ danger: true, disabled: !tightenTypes.length }"
        @ok="executeTighten"
      >
        <div class="modal-copy">
          <div class="modal-safety"><SafetyOutlined /><strong>已选 {{ selectedSheets.length }} 个文件</strong></div>
          <p>选择要关闭的共享类型。已明确添加的用户和群组邮箱不会被移除，所有者权限也不会改变。</p>
          <a-checkbox-group v-model:value="tightenTypes" class="tighten-options">
            <a-checkbox value="anyone">
              <span><strong>关闭“知道链接的任何人”</strong><small>移除 anyone 类型权限</small></span>
            </a-checkbox>
            <a-checkbox value="domain">
              <span><strong>关闭整个 Workspace 域共享</strong><small>移除 domain 类型权限；继承项按下方选项处理</small></span>
            </a-checkbox>
          </a-checkbox-group>
          <a-checkbox v-model:checked="limitedAccessOnInherited" class="limited-access-option">
            <span>
              <strong>继承错误时自动启用文件夹 limited access</strong>
              <small>
                高风险：将作用于约 {{ selectedParentFolderCount }} 个直属父文件夹，并阻断这些文件夹的所有继承访问，
                包括继承的具体用户和群组；直接添加的权限与所有者不受影响。
              </small>
            </span>
          </a-checkbox>
          <a-alert
            :type="limitedAccessOnInherited ? 'error' : 'warning'"
            show-icon
            :message="`预计移除 ${selectedRiskCount} 项直接共享权限；所有明确添加的账号都会保留`"
            :description="
              selectedInheritedRiskCount
                ? limitedAccessOnInherited
                  ? `${selectedInheritedRiskCount} 项继承权限将通过直属父文件夹的 limited access 处理；这可能让其他文件也失去继承访问。`
                  : `${selectedInheritedRiskCount} 项继承权限会跳过；勾选上方高风险选项可自动尝试处理。`
                : undefined
            "
          />
        </div>
      </a-modal>

      <div v-if="progress.visible && (granting || openSharingRunning || tightenRunning || trashRunning)" class="progress-dock">
        <div class="progress-copy">
          <CloudSyncOutlined spin />
          <div><strong>{{ progress.label }}</strong><span>{{ progress.current }} / {{ progress.total }}</span></div>
        </div>
        <a-progress
          :percent="progress.total ? Math.round((progress.current / progress.total) * 100) : 0"
          :show-info="false"
          size="small"
        />
      </div>
    </div>
  </a-config-provider>
</template>
