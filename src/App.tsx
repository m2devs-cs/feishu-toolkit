import React, { useCallback, useMemo, useState } from 'react';
import { bitable } from '@lark-base-open/js-sdk';

function cellToString(val: any): string {
  if (val == null) return '';
  if (typeof val === 'string') return val;
  if (Array.isArray(val)) {
    return (val as any[])
      .map((v: any) =>
        v && typeof v.text === 'string'
          ? v.text
          : typeof v === 'string'
          ? v
          : '',
      )
      .join('');
  }
  if (typeof val === 'object') {
    const v: any = val as any;
    if (typeof v.link === 'string') return v.link;
    if (typeof v.text === 'string') return v.text;
  }
  return '';
}

export default function App() {
  const [result, setResult] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [status, setStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [copied, setCopied] = useState<boolean>(false);
  const [bulkLoading, setBulkLoading] = useState<boolean>(false);
  const [bulkProgress, setBulkProgress] = useState<string>('');
  const [bulkSummary, setBulkSummary] = useState<string>('');

  const styles = useMemo(
    () => ({
      container: {
        padding: 16,
        fontFamily:
          "-apple-system, BlinkMacSystemFont, 'Helvetica Neue', Tahoma, 'PingFang TC', 'Microsoft Jhenghei', Arial, sans-serif",
        color: '#1f2329',
      } as React.CSSProperties,
      header: {
        fontSize: 16,
        fontWeight: 600,
        marginBottom: 12,
      } as React.CSSProperties,
      card: {
        border: '1px solid #e5e6eb',
        borderRadius: 8,
        padding: 16,
        background: '#fff',
      } as React.CSSProperties,
      desc: { fontSize: 13, color: '#646a73', marginBottom: 12 } as React.CSSProperties,
      btnPrimary: {
        display: 'inline-block',
        width: '100%',
        padding: '10px 16px',
        borderRadius: 6,
        border: '1px solid transparent',
        background: '#3370ff',
        color: '#fff',
        cursor: 'pointer',
        fontSize: 14,
      } as React.CSSProperties,
      btnSecondary: {
        display: 'inline-block',
        width: '100%',
        padding: '10px 16px',
        borderRadius: 6,
        border: '1px solid #e5e6eb',
        background: '#fff',
        color: '#1f2329',
        cursor: 'pointer',
        fontSize: 14,
      } as React.CSSProperties,
      gap: { height: 10 } as React.CSSProperties,
      success: { color: '#3370ff', textDecoration: 'underline', cursor: 'pointer' } as React.CSSProperties,
      error: { color: '#f53f3f' } as React.CSSProperties,
      small: { fontSize: 12, color: '#8f959e', marginLeft: 8 } as React.CSSProperties,
      progress: { marginTop: 8, fontSize: 13, color: '#1f2329' } as React.CSSProperties,
    }),
    [],
  );

  const generateShortUrl = useCallback(async () => {
    setLoading(true);
    setResult('');
    setStatus('idle');
    try {
      const selection = await bitable.base.getSelection();
      if (!selection.tableId || !selection.recordId) {
        throw new Error('no selection');
      }
      const table = await bitable.base.getTableById(selection.tableId);
      const longField = await table.getFieldByName('網址');
      const shortField = await table.getFieldByName('短網址');
      const val = await longField.getValue(selection.recordId);
      const longUrl = cellToString(val);
      if (!longUrl) {
        setResult('找不到網址，請確認「網址」欄位有填入內容');
        setLoading(false);
        setStatus('error');
        return;
      }
      const resp = await fetch(
        `https://tinyurl.com/api-create.php?url=${encodeURIComponent(longUrl)}`,
      );
      const text = await resp.text();
      if (!resp.ok || !text || /^Error/i.test(text)) {
        throw new Error('tinyurl error');
      }
      const shortUrl = text;
      await shortField.setValue(selection.recordId, shortUrl);
      setResult(shortUrl);
      setStatus('success');
    } catch {
      setResult('生成失敗，請確認網址是否正確');
      setStatus('error');
    } finally {
      setLoading(false);
    }
  }, []);

  const copyShortUrl = useCallback(async () => {
    if (status !== 'success' || !result) return;
    try {
      await navigator.clipboard.writeText(result);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch {}
  }, [status, result]);

  const bulkGenerateEmptyShortUrls = useCallback(async () => {
    setBulkLoading(true);
    setBulkProgress('');
    setBulkSummary('');
    try {
      const table = await bitable.base.getActiveTable();
      const longField = await table.getFieldByName('網址');
      const shortField = await table.getFieldByName('短網址');
      const recordIds = await table.getRecordIdList();

      const candidates: Array<{ recordId: string; longUrl: string }> = [];
      for (const recordId of recordIds) {
        const longVal = cellToString(await longField.getValue(recordId));
        const shortVal = cellToString(await shortField.getValue(recordId));
        if (longVal && !shortVal) {
          candidates.push({ recordId, longUrl: longVal });
        }
      }

      let success = 0;
      let failure = 0;
      const total = candidates.length;
      let index = 0;
      for (const item of candidates) {
        index += 1;
        setBulkProgress(`正在處理第 ${index} / 共 ${total} 行...`);
        try {
          const resp = await fetch(
            `https://tinyurl.com/api-create.php?url=${encodeURIComponent(item.longUrl)}`,
          );
          const text = await resp.text();
          if (!resp.ok || !text || /^Error/i.test(text)) {
            throw new Error('tinyurl error');
          }
          await shortField.setValue(item.recordId, text);
          success += 1;
        } catch {
          failure += 1;
        }
      }

      if (failure === 0) {
        setBulkSummary(`完成，已生成 ${success} 條短網址`);
      } else {
        setBulkSummary(`完成，成功 ${success} 條，失敗 ${failure} 條`);
      }
    } catch {
      setBulkSummary('處理失敗，請稍後重試');
    } finally {
      setBulkLoading(false);
      setBulkProgress('');
    }
  }, []);

  return (
    <div style={styles.container}>
      <div style={styles.header}>🔗 短鏈生成器</div>
      <div style={styles.card}>
        <p style={styles.desc}>
          點「生成短網址」處理當前選中的行；點「批量生成」自動填滿所有空白的短網址欄位
        </p>
        <button
          style={styles.btnPrimary}
          disabled={bulkLoading || loading}
          onClick={bulkGenerateEmptyShortUrls}
        >
          {bulkLoading ? '生成中...' : '批量生成空白短網址'}
        </button>
        <div style={styles.gap} />
        <button
          style={styles.btnSecondary}
          disabled={loading || bulkLoading}
          onClick={generateShortUrl}
        >
          生成短網址
        </button>
        {(bulkLoading || bulkSummary) && (
          <div style={styles.progress}>
            {bulkLoading ? bulkProgress : bulkSummary}
          </div>
        )}
        {Boolean(result) && (
          <div style={{ marginTop: 8 }}>
            {status === 'success' ? (
              <div>
                <span style={styles.success} onClick={copyShortUrl}>
                  {result}
                </span>
                {copied && <span style={styles.small}>已複製</span>}
              </div>
            ) : (
              <div style={styles.error}>{result}</div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
