import {workbook} from '@sebbebroman/exceljs';
import * as main from '@sebbebroman/exceljs';
import {sheetProtection, type SheetProtection} from '@sebbebroman/exceljs/protection';
const model: SheetProtection = sheetProtection('secret', {spinCount: 1, selectLockedCells: false});
await workbook().sheet('Locked').protect(model).writeBuffer();
workbook().sheet('NoPassword').protect({sheet: true});
// @ts-expect-error Password hashing must be an explicit opt-in import.
void main.sheetProtection;
// @ts-expect-error The core builder accepts prepared protection, not a password.
workbook().sheet('Old').protect('secret');
