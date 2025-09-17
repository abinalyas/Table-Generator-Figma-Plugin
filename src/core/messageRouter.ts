import { handleTableCreation, handleTableUpdate } from '../handlers/tableHandlers';
import { handleTableScan, handleRequestComponentInfo } from '../handlers/scanHandlers';
import { handleAITableGeneration, handleWatsonxDataGeneration } from '../handlers/aiHandlers';

export function initMessageRouter(): void {
  figma.ui.onmessage = async (msg: any) => {
    if (msg.type === 'resize') {
      const { size } = msg; let { w, h } = size;
      // Constrain and forward to setSize to avoid UI breakage
      w = Math.max(320, Math.min(1200, Math.floor(w)));
      h = Math.max(400, Math.min(1200, Math.floor(h)));
      figma.ui.resize(w, h);
      return;
    }

    if (msg.type === 'resize-ui') {
      const { width, height } = msg;
      const w = Math.max(320, Math.min(1200, Math.floor(width)));
      const h = Math.max(400, Math.min(1200, Math.floor(height)));
      figma.ui.resize(w, h);
      return;
    }

    if (msg.type === 'scan-selected-table') {
      // Delegate to selection flow already handled by selectionchange - no-op here
      return;
    }

    if (msg.type === 'generate-table-with-ai') {
      await handleAITableGeneration(msg);
      return;
    }

    if (msg.type === 'generate-watsonx-data') {
      await handleWatsonxDataGeneration(msg);
      return;
    }

    if (msg.type === 'scan-table') {
      await handleTableScan(msg);
      return;
    }

    if (msg.type === 'request-component-info') {
      await handleRequestComponentInfo(msg);
      return;
    }

    if (msg.type === 'create-table-from-scan') {
      await handleTableCreation(msg);
      return;
    }

    if (msg.type === 'update-table') {
      await handleTableUpdate(msg);
      return;
    }
  };
}


