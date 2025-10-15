export type TableSettings = {
  columns: number;
  rows: number;
  includeHeader?: boolean;
  includeFooter?: boolean;
  includeToolbar?: boolean;
  includeSelectable?: boolean;
  includeExpandable?: boolean;
  cellProperties?: any;
};

const FLAG_KEY = 'isGeneratedTable';
const SETTINGS_KEY = 'tableSettings';

export function markAsGeneratedTable(node: BaseNode & PluginDataMixin) {
  node.setPluginData(FLAG_KEY, 'true');
}

export function writeTableSettings(node: BaseNode & PluginDataMixin, settings: TableSettings) {
  try {
    node.setPluginData(SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) {
    // Best-effort only; do not throw in plugin runtime
    console.error('Failed to write table settings', e);
  }
}

export function readTableSettings(node: BaseNode & PluginDataMixin): TableSettings | null {
  try {
    const raw = node.getPluginData(SETTINGS_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as TableSettings;
  } catch (e) {
    console.error('Failed to read/parse table settings', e);
    return null;
  }
}

export function isGeneratedTable(node: BaseNode & PluginDataMixin): boolean {
  try {
    return node.getPluginData(FLAG_KEY) === 'true';
  } catch {
    return false;
  }
}


