
import { TableService } from 'azure-storage';

export interface Environment {
    ewsUrl: string;
    ewsAuthType: string;
    ewsToken?: string;
    ewsUser?: string;
    ewsPassword?: string;
    tableService: TableService;
    logId: number;
    logCount: number;
}