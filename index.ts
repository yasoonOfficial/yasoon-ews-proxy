import { Monkey } from './extensions/Monkey';

//Fix for https://github.com/gautamsi/ews-javascript-api/pull/219
new Monkey().patch();

export * from './proxy/create-calendar';
export * from './proxy/create-event';
export * from './proxy/create-wunderbar-link';
export * from './proxy/delete-calendar';
export * from './proxy/delete-event';
export * from './proxy/get-autodiscover-data';
export * from './proxy/get-calendars';
export * from './proxy/get-categories';
export * from './proxy/get-events';
export * from './proxy/get-free-busy-events';
export * from './proxy/get-permissions';
export * from './proxy/get-publicfolder-mailbox';
export * from './proxy/get-user-image';
export * from './proxy/get-user';
export * from './proxy/search-user';
export * from './proxy/mapper';
export * from './proxy/helper';
export * from './proxy/update-event';

export * from './model/constants';
export * from './model/office';
export * from './model/proxy';