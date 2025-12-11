'use client';

import { useState } from 'react';

export default function Home() {
  const [templateBase64, setTemplateBase64] = useState<string>('');
  const [placeholderData, setPlaceholderData] = useState<string>('');
  const [outputFormat, setOutputFormat] = useState<'excel' | 'pdf'>('excel');
  const [isGenerating, setIsGenerating] = useState(false);
  const [placeholders, setPlaceholders] = useState<string[]>([]);
  const [templateInfo, setTemplateInfo] = useState<any>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      const base64 = event.target?.result as string;
      const base64Data = base64.split(',')[1];
      setTemplateBase64(base64Data);

      // プレースホルダーを検出
      try {
        const response = await fetch('/api/template/placeholders', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ templateBase64: base64Data }),
        });

        const result = await response.json();
        if (result.success) {
          setPlaceholders(result.placeholders);
        }
      } catch (error) {
        console.error('Error finding placeholders:', error);
      }

      // テンプレート情報を取得
      try {
        const response = await fetch('/api/template/info', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ templateBase64: base64Data }),
        });

        const result = await response.json();
        if (result.success) {
          setTemplateInfo(result.templateInfo);
        }
      } catch (error) {
        console.error('Error getting template info:', error);
      }
    };

    reader.readAsDataURL(file);
  };

  const handleGenerate = async () => {
    if (!templateBase64) {
      alert('テンプレートファイルをアップロードしてください');
      return;
    }

    if (!placeholderData) {
      alert('プレースホルダーデータを入力してください');
      return;
    }

    try {
      setIsGenerating(true);
      const data = JSON.parse(placeholderData);

      const endpoint =
        outputFormat === 'excel' ? '/api/generate/excel' : '/api/generate/pdf';

      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          templateBase64,
          data,
          options: {
            orientation: 'portrait',
            format: 'a4',
          },
        }),
      });

      const result = await response.json();

      if (result.success) {
        // ダウンロード
        const blob = base64ToBlob(result.data, result.mimeType);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `report.${outputFormat === 'excel' ? 'xlsx' : 'pdf'}`;
        a.click();
        URL.revokeObjectURL(url);
      } else {
        // LibreOfficeがインストールされていない場合の詳細なエラーメッセージ
        if (result.installInstructions) {
          const instructions = result.installInstructions.windows || result.details;
          alert(
            `${result.error}\n\n` +
            `${result.details}\n\n` +
            `インストール方法:\n${instructions}`
          );
        } else {
          alert(`生成に失敗しました: ${result.error}\n\n詳細: ${result.details || ''}`);
        }
      }
    } catch (error) {
      console.error('Error generating report:', error);
      alert('生成中にエラーが発生しました');
    } finally {
      setIsGenerating(false);
    }
  };

  const base64ToBlob = (base64: string, mimeType: string): Blob => {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: mimeType });
  };

  const generateSampleData = () => {
    const sample: any = {};
    placeholders.forEach((placeholder) => {
      sample[placeholder] = `サンプル値_${placeholder}`;
    });
    setPlaceholderData(JSON.stringify(sample, null, 2));
  };

  return (
    <div className="min-h-screen bg-gray-50 py-8 px-4">
      <div className="max-w-4xl mx-auto">
        <h1 className="text-3xl font-bold text-gray-900 mb-8">
          エクセル帳票生成システム
        </h1>

        <div className="bg-white rounded-lg shadow p-6 mb-6">
          <h2 className="text-xl font-semibold mb-4">1. テンプレートアップロード</h2>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
          />

          {templateInfo && (
            <div className="mt-4 p-4 bg-gray-50 rounded">
              <h3 className="font-semibold mb-2">テンプレート情報:</h3>
              <p>シート数: {templateInfo.sheetCount}</p>
              <ul className="list-disc list-inside">
                {templateInfo.sheets.map((sheet: any) => (
                  <li key={sheet.id}>
                    {sheet.name} ({sheet.rowCount}行 × {sheet.columnCount}列)
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>

        {placeholders.length > 0 && (
          <div className="bg-white rounded-lg shadow p-6 mb-6">
            <h2 className="text-xl font-semibold mb-4">
              2. 検出されたプレースホルダー
            </h2>
            <div className="flex flex-wrap gap-2">
              {placeholders.map((placeholder) => (
                <span
                  key={placeholder}
                  className="px-3 py-1 bg-blue-100 text-blue-800 rounded-full text-sm"
                >
                  {`{{${placeholder}}}`}
                </span>
              ))}
            </div>
            <button
              onClick={generateSampleData}
              className="mt-4 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
            >
              サンプルデータを生成
            </button>
          </div>
        )}

        <div className="bg-white rounded-lg shadow p-6 mb-6">
          <h2 className="text-xl font-semibold mb-4">3. プレースホルダーデータ入力</h2>
          <textarea
            value={placeholderData}
            onChange={(e) => setPlaceholderData(e.target.value)}
            placeholder='{"ClassA": "値1", "ClassB": "値2"}'
            className="w-full h-48 p-3 border border-gray-300 rounded-lg font-mono text-sm"
          />
          <p className="text-sm text-gray-500 mt-2">
            JSON形式でプレースホルダーの値を入力してください
          </p>
        </div>

        <div className="bg-white rounded-lg shadow p-6 mb-6">
          <h2 className="text-xl font-semibold mb-4">4. 出力形式選択</h2>
          <div className="flex gap-4">
            <label className="flex items-center">
              <input
                type="radio"
                value="excel"
                checked={outputFormat === 'excel'}
                onChange={(e) => setOutputFormat(e.target.value as 'excel' | 'pdf')}
                className="mr-2"
              />
              <span>Excelファイル (.xlsx)</span>
            </label>
            <label className="flex items-center">
              <input
                type="radio"
                value="pdf"
                checked={outputFormat === 'pdf'}
                onChange={(e) => setOutputFormat(e.target.value as 'excel' | 'pdf')}
                className="mr-2"
              />
              <span>PDFファイル (.pdf)</span>
            </label>
          </div>
        </div>

        <div className="bg-white rounded-lg shadow p-6">
          <button
            onClick={handleGenerate}
            disabled={isGenerating || !templateBase64}
            className="w-full py-3 px-4 bg-blue-600 text-white rounded-lg font-semibold hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            {isGenerating ? '生成中...' : '帳票を生成'}
          </button>
        </div>

        <div className="mt-8 bg-blue-50 rounded-lg p-6">
          <h3 className="font-semibold text-blue-900 mb-2">使い方:</h3>
          <ol className="list-decimal list-inside space-y-2 text-blue-800">
            <li>
              プリザンターにアップロード済みのエクセルファイルをダウンロードし、アップロード
            </li>
            <li>
              テンプレート内の<code className="bg-blue-100 px-1 rounded">{'{{ClassA}}'}</code>
              形式のプレースホルダーが自動検出されます
            </li>
            <li>JSON形式でプレースホルダーに対応する値を入力</li>
            <li>出力形式（ExcelまたはPDF）を選択</li>
            <li>「帳票を生成」ボタンでダウンロード</li>
          </ol>
        </div>
      </div>
    </div>
  );
}
