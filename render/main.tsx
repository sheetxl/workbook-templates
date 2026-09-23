/**
 * Renders one template for its thumbnail: `?src=<category>/<name>.xlsx`.
 *
 * The studio is the same renderer the gallery opens the template in, so the thumbnail shows what
 * the user will get. Its chrome is hidden; only the grid is photographed.
 *
 * `window.__thumbnail` reports the outcome to `config/thumbnails.ts`: `'ready'` once the grid has
 * painted, or an error message.
 */
import { createRoot } from 'react-dom/client';

import { Studio } from '@sheetxl/studio-mui';

declare global {
  interface Window { __thumbnail?: 'ready' | { error: string } }
}

const src = new URLSearchParams(window.location.search).get('src');

const App = () => {
  if (!src) {
    window.__thumbnail = { error: 'no ?src given' };
    return null;
  }
  return (
    <Studio
      workbook={{ source: new URL(src, window.location.href).href }}
      showFormulaBar={false}
      showStatusBar={false}
      renderToolbar={() => <></>}
      propsWorkbook={{ showTabs: false, showHorizontalScrollbar: false, showVerticalScrollbar: false }}
      square
      style={{ position: 'absolute', inset: 0 }}
      onWorkbookLoad={() => {
        // The load event fires one frame before the grid paints.
        requestAnimationFrame(() => requestAnimationFrame(() => { window.__thumbnail = 'ready'; }));
      }}
      onDocumentError={(error: unknown) => {
        window.__thumbnail = { error: error instanceof Error ? error.message : String(error) };
      }}
    />
  );
};

createRoot(document.getElementById('root')!).render(<App/>);
