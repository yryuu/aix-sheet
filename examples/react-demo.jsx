/**
 * react-demo.jsx — How to embed aix-sheet's viewer in a React app.
 *
 * Requires React 17+. Use with a bundler (Vite, webpack, etc.).
 *
 *   npm install aix-sheet react react-dom
 */
import React, { useMemo, useRef, useState } from 'react';
import { Sheet } from 'aix-sheet';
import { SheetViewReact } from 'aix-sheet/react';
import 'aix-sheet/viewer/styles.css';

export default function App() {
  const sheet = useMemo(() => {
    const s = new Sheet('Sales');
    s.write('A1:C1', ['商品', '価格', '在庫']);
    s.style('A1:C1', { bold: true, bgColor: '#217346', color: '#ffffff', align: 'center' });
    s.write('A2:A4', ['りんご', 'みかん', 'ぶどう']);
    s.write('B2:B4', [120, 80, 350]);
    s.write('C2:C4', [50, 100, 20]);
    s.write('A5', '合計');
    s.write('B5', '=SUM(B2:B4)');
    s.write('C5', '=SUM(C2:C4)');
    return s;
  }, []);

  const viewRef = useRef(null);
  const [lastChange, setLastChange] = useState(null);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      <div style={{ padding: 8, background: '#eee', display: 'flex', gap: 12 }}>
        <button onClick={() => viewRef.current?.focusCell('B5')}>B5 にフォーカス</button>
        <button onClick={() => {
          const s = viewRef.current?.getSheet();
          if (s) {
            s.write('A6', 'New');
            viewRef.current.refresh();
          }
        }}>外部から A6 に "New" を書き込み</button>
        <span style={{ marginLeft: 'auto' }}>
          {lastChange ? `変更: ${lastChange.ref} → ${lastChange.newValue}` : ''}
        </span>
      </div>
      <SheetViewReact
        ref={viewRef}
        sheet={sheet}
        onChange={setLastChange}
        style={{ flex: 1 }}
      />
    </div>
  );
}
