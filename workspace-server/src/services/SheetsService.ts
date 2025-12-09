/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { google, sheets_v4, drive_v3 } from 'googleapis';
import { AuthManager } from '../auth/AuthManager';
import { logToFile } from '../utils/logger';
import { extractDocId } from '../utils/IdUtils';
import { gaxiosOptions } from '../utils/GaxiosConfig';
import { buildDriveSearchQuery, MIME_TYPES } from '../utils/DriveQueryBuilder';

export class SheetsService {
    constructor(private authManager: AuthManager) {
    }

    private async getSheetsClient(): Promise<sheets_v4.Sheets> {
        const auth = await this.authManager.getAuthenticatedClient();
        const options = { ...gaxiosOptions, auth };
        return google.sheets({ version: 'v4', ...options });
    }

    private async getDriveClient(): Promise<drive_v3.Drive> {
        const auth = await this.authManager.getAuthenticatedClient();
        const options = { ...gaxiosOptions, auth };
        return google.drive({ version: 'v3', ...options });
    }

    public getText = async ({ spreadsheetId, format = 'text' }: { spreadsheetId: string, format?: 'text' | 'csv' | 'json' }) => {
        logToFile(`[SheetsService] Starting getText for spreadsheet: ${spreadsheetId} with format: ${format}`);
        try {
            const id = extractDocId(spreadsheetId) || spreadsheetId;

            const sheets = await this.getSheetsClient();
            // Get spreadsheet metadata
            const spreadsheet = await sheets.spreadsheets.get({
                spreadsheetId: id,
                includeGridData: false,
            });

            let content = '';
            const jsonData: Record<string, any[][]> = {};

            // Add spreadsheet title (except for JSON format)
            if (spreadsheet.data.properties?.title && format !== 'json') {
                content += `Spreadsheet Title: ${spreadsheet.data.properties.title}\n\n`;
            }

            // Get all sheet names
            const sheetNames = spreadsheet.data.sheets?.map(sheet => sheet.properties?.title) || [];

            // Get data from all sheets
            for (const sheetName of sheetNames) {
                if (!sheetName) continue;

                try {
                    const response = await sheets.spreadsheets.values.get({
                        spreadsheetId: id,
                        range: `'${sheetName}'`,
                    });

                    const values = response.data.values || [];

                    if (format === 'json') {
                        // Collect data for JSON structure
                        jsonData[sheetName] = values;
                    } else {
                        // Add sheet name as context
                        content += `Sheet Name: ${sheetName}\n`;

                        if (values.length === 0) {
                            content += '(Empty sheet)\n';
                        } else {
                            // Process each row
                            values.forEach((row) => {
                                if (format === 'csv') {
                                    // Convert to CSV format
                                    const csvRow = row.map(cell => {
                                        // Escape quotes and wrap in quotes if contains comma or quotes
                                        const cellStr = String(cell || '');
                                        if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                                            return `"${cellStr.replace(/"/g, '""')}"`;
                                        }
                                        return cellStr;
                                    }).join(',');
                                    content += csvRow + '\n';
                                } else {
                                    // Plain text format with pipe separators for readability
                                    content += row.map(cell => cell || '').join(' | ') + '\n';
                                }
                            });
                        }
                        content += '\n';
                    }
                } catch (sheetError) {
                    logToFile(`[SheetsService] Error reading sheet ${sheetName}: ${sheetError}`);
                    if (format === 'json') {
                        // For JSON format, we'll skip sheets with errors
                        logToFile(`[SheetsService] Skipping sheet ${sheetName} in JSON output due to error`);
                    } else {
                        content += `Sheet Name: ${sheetName}\n(Error reading sheet)\n\n`;
                    }
                }
            }

            if (format === 'json') {
                // Generate clean JSON output from collected data
                content = JSON.stringify(jsonData, null, 2);
            }

            logToFile(`[SheetsService] Finished getText for spreadsheet: ${id}`);
            return {
                content: [{
                    type: "text" as const,
                    text: content.trim()
                }]
            };
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logToFile(`[SheetsService] Error during sheets.getText: ${errorMessage}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({ error: errorMessage })
                }]
            };
        }
    }

    public getRange = async ({ spreadsheetId, range }: { spreadsheetId: string, range: string }) => {
        logToFile(`[SheetsService] Starting getRange for spreadsheet: ${spreadsheetId}, range: ${range}`);
        try {
            const id = extractDocId(spreadsheetId) || spreadsheetId;

            const sheets = await this.getSheetsClient();
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: id,
                range: range,
            });

            const values = response.data.values || [];

            logToFile(`[SheetsService] Finished getRange for spreadsheet: ${id}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({
                        range: response.data.range,
                        values: values
                    })
                }]
            };
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logToFile(`[SheetsService] Error during sheets.getRange: ${errorMessage}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({ error: errorMessage })
                }]
            };
        }
    }

    public find = async ({ query, pageToken, pageSize = 10 }: { query: string, pageToken?: string, pageSize?: number }) => {
        logToFile(`[SheetsService] Searching for spreadsheets with query: ${query}`);
        try {
            const q = buildDriveSearchQuery(MIME_TYPES.SPREADSHEET, query);
            logToFile(`[SheetsService] Executing Drive API query: ${q}`);

            const drive = await this.getDriveClient();
            const res = await drive.files.list({
                pageSize: pageSize,
                fields: 'nextPageToken, files(id, name)',
                q: q,
                pageToken: pageToken,
            });

            const files = res.data.files || [];
            const nextPageToken = res.data.nextPageToken;

            logToFile(`[SheetsService] Found ${files.length} spreadsheets.`);

            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({
                        files: files,
                        nextPageToken: nextPageToken
                    })
                }]
            };
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logToFile(`[SheetsService] Error during sheets.find: ${errorMessage}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({ error: errorMessage })
                }]
            };
        }
    }

    public getMetadata = async ({ spreadsheetId }: { spreadsheetId: string }) => {
        logToFile(`[SheetsService] Starting getMetadata for spreadsheet: ${spreadsheetId}`);
        try {
            const id = extractDocId(spreadsheetId) || spreadsheetId;

            const sheets = await this.getSheetsClient();
            const spreadsheet = await sheets.spreadsheets.get({
                spreadsheetId: id,
                includeGridData: false,
            });

            const metadata = {
                spreadsheetId: spreadsheet.data.spreadsheetId,
                title: spreadsheet.data.properties?.title,
                sheets: spreadsheet.data.sheets?.map(sheet => ({
                    sheetId: sheet.properties?.sheetId,
                    title: sheet.properties?.title,
                    index: sheet.properties?.index,
                    rowCount: sheet.properties?.gridProperties?.rowCount,
                    columnCount: sheet.properties?.gridProperties?.columnCount,
                })),
                locale: spreadsheet.data.properties?.locale,
                timeZone: spreadsheet.data.properties?.timeZone,
            };

            logToFile(`[SheetsService] Finished getMetadata for spreadsheet: ${id}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify(metadata)
                }]
            };
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logToFile(`[SheetsService] Error during sheets.getMetadata: ${errorMessage}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({ error: errorMessage })
                }]
            };
        }
    }

    public pasteCSVData = async ({ spreadsheetId, csvData, title, startRow = 0, startColumn = 0 }: {
        spreadsheetId: string,
        csvData: string,
        title: string,
        startRow?: number,
        startColumn?: number
    }) => {
        logToFile(`[SheetsService] Starting pasteCSVData for spreadsheet: ${spreadsheetId}, sheet: ${title}`);
        try {
            const id = extractDocId(spreadsheetId) || spreadsheetId;
            const sheets = await this.getSheetsClient();

            // 1. Check if sheet exists
            const spreadsheet = await sheets.spreadsheets.get({
                spreadsheetId: id,
                includeGridData: false,
            });

            const existingSheet = spreadsheet.data.sheets?.find(
                (s: sheets_v4.Schema$Sheet) => s.properties?.title === title
            );

            let targetSheetId: number;
            const requests: sheets_v4.Schema$Request[] = [];

            if (existingSheet) {
                // Sheet exists, use its ID
                logToFile(`[SheetsService] Sheet "${title}" exists (ID: ${existingSheet.properties?.sheetId})`);
                targetSheetId = existingSheet.properties?.sheetId || 0;
            } else {
                // Sheet doesn't exist, create it
                // Generate a random ID to avoid collisions (though unlikely with one user)
                // Range is safe for 32-bit integer
                targetSheetId = Math.floor(Math.random() * 2147483647);
                logToFile(`[SheetsService] Creating new sheet "${title}" with ID: ${targetSheetId}`);

                requests.push({
                    addSheet: {
                        properties: {
                            title: title,
                            sheetId: targetSheetId
                        }
                    }
                });
            }

            // 2. Add pasteData request
            requests.push({
                pasteData: {
                    coordinate: {
                        sheetId: targetSheetId,
                        rowIndex: startRow,
                        columnIndex: startColumn
                    },
                    data: csvData,
                    type: "PASTE_NORMAL",
                    delimiter: ","
                }
            });

            // 3. Execute batchUpdate
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: id,
                requestBody: {
                    requests: requests
                }
            });

            logToFile(`[SheetsService] Finished pasteCSVData for spreadsheet: ${id}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({
                        status: "success",
                        message: existingSheet
                            ? `Updated existing sheet "${title}"`
                            : `Created new sheet "${title}" and populated with data`,
                        sheetId: targetSheetId
                    })
                }]
            };

        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logToFile(`[SheetsService] Error during sheets.pasteCSVData: ${errorMessage}`);
            return {
                content: [{
                    type: "text" as const,
                    text: JSON.stringify({ error: errorMessage })
                }]
            };
        }
    }
}
