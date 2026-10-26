import { describe, expect, it } from 'vitest'
import { countInvalidEmails, escapeCsvCell, parseEmails, parseGoogleFileIds } from './parsers'

describe('input parsers', () => {
  it('extracts and deduplicates Google file IDs', () => {
    const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz'
    expect(parseGoogleFileIds(`https://docs.google.com/spreadsheets/d/${id}/edit\n${id}`)).toEqual([id])
  })

  it('extracts Docs, Slides, Drive file and folder IDs', () => {
    const documentId = '1DocumentAbCdEfGhIjKlMnOpQrSt'
    const presentationId = '1PresentationAbCdEfGhIjKlMnOp'
    const fileId = '1DriveFileAbCdEfGhIjKlMnOpQrSt'
    const folderId = '1DriveFolderAbCdEfGhIjKlMnOpQr'
    expect(
      parseGoogleFileIds(
        [
          `https://docs.google.com/document/d/${documentId}/edit`,
          `https://docs.google.com/presentation/d/${presentationId}/edit`,
          `https://drive.google.com/file/d/${fileId}/view`,
          `https://drive.google.com/drive/folders/${folderId}`,
        ].join('\n'),
      ),
    ).toEqual([documentId, presentationId, fileId, folderId])
  })

  it('normalizes and deduplicates emails', () => {
    expect(parseEmails('Alice@Example.com, alice@example.com\nbob@example.com')).toEqual([
      'alice@example.com',
      'bob@example.com',
    ])
  })

  it('preserves repeated items when paired mode disables deduplication', () => {
    const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz'
    expect(parseGoogleFileIds(`${id}\n${id}`, false)).toEqual([id, id])
    expect(parseEmails('same@example.com\nsame@example.com', false)).toEqual(['same@example.com', 'same@example.com'])
  })

  it('counts invalid email tokens', () => {
    expect(countInvalidEmails('ok@example.com, not-an-email')).toBe(1)
  })

  it('neutralizes spreadsheet formulas when exporting CSV cells', () => {
    expect(escapeCsvCell('=HYPERLINK("https://evil.example")')).toBe(
      '"\'=HYPERLINK(""https://evil.example"")"',
    )
    expect(escapeCsvCell('  +SUM(1,1)')).toBe('"\'  +SUM(1,1)"')
    expect(escapeCsvCell('\t@malicious')).toBe('"\'\t@malicious"')
    expect(escapeCsvCell('普通文件名')).toBe('"普通文件名"')
  })
})
