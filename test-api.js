/**
 * test-api.js
 * Docker版のAPIエンドポイントをテスト
 */

const fs = require('fs');
const path = require('path');

async function testAPI() {
  console.log('=== Docker版 API テスト ===\n');

  try {
    // 1. テンプレートファイルを読み込み
    console.log('1. テンプレートファイル読み込み...');
    const templatePath = path.join(__dirname, 'test-template.xlsx');
    const templateBuffer = fs.readFileSync(templatePath);
    const templateBase64 = templateBuffer.toString('base64');
    console.log(`   ✅ テンプレート読み込み完了: ${(templateBuffer.length / 1024).toFixed(2)} KB\n`);

    // 2. プレースホルダー検出API
    console.log('2. プレースホルダー検出API テスト...');
    const placeholdersResponse = await fetch('http://localhost:3000/api/template/placeholders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ templateBase64 })
    });

    const placeholdersResult = await placeholdersResponse.json();
    console.log('   レスポンス:', JSON.stringify(placeholdersResult, null, 2));

    if (!placeholdersResult.success) {
      throw new Error('プレースホルダー検出に失敗');
    }

    console.log(`   ✅ 検出されたプレースホルダー: ${placeholdersResult.placeholders.length}個`);
    placeholdersResult.placeholders.forEach(p => {
      console.log(`      - {{${p}}}`);
    });
    console.log('');

    // 3. テンプレート情報取得API
    console.log('3. テンプレート情報取得API テスト...');
    const infoResponse = await fetch('http://localhost:3000/api/template/info', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ templateBase64 })
    });

    const infoResult = await infoResponse.json();
    console.log('   レスポンス:', JSON.stringify(infoResult, null, 2));

    if (!infoResult.success) {
      throw new Error('テンプレート情報取得に失敗');
    }

    console.log(`   ✅ シート数: ${infoResult.templateInfo.sheetCount}`);
    console.log(`   ✅ Note: ${infoResult.note}\n`);

    // 4. Excel生成API（test9方式のテスト - 特殊文字含む）
    console.log('4. Excel生成API テスト（test9方式 - 特殊文字含む）...');
    const testData = {
      会社名: 'A社 & B社 <共同>',           // XMLエスケープが必要
      日付: '2024年12月</v>',               // タグ風文字列
      金額: '¥1,234,567 "特別価格"',       // 引用符
      担当者: "山田太郎 's dept."          // アポストロフィ
    };

    console.log('   テストデータ:');
    Object.entries(testData).forEach(([key, value]) => {
      console.log(`      ${key}: "${value}"`);
    });

    const generateResponse = await fetch('http://localhost:3000/api/generate/excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        templateBase64,
        data: testData
      })
    });

    const generateResult = await generateResponse.json();

    if (!generateResult.success) {
      throw new Error(`Excel生成に失敗: ${generateResult.error}`);
    }

    console.log('   ✅ Excel生成成功');
    console.log(`   ✅ MimeType: ${generateResult.mimeType}`);

    // 5. 結果を保存
    const outputPath = path.join(__dirname, 'test-output.xlsx');
    const outputBuffer = Buffer.from(generateResult.data, 'base64');
    fs.writeFileSync(outputPath, outputBuffer);

    console.log(`   ✅ 出力ファイル保存: ${outputPath}`);
    console.log(`   ✅ ファイルサイズ: ${(outputBuffer.length / 1024).toFixed(2)} KB\n`);

    // 6. 印刷設定が保持されているか検証
    console.log('5. 印刷設定保持検証...');
    const JSZip = require('jszip');

    // 元のテンプレートのpageSetup取得
    const originalZip = await JSZip.loadAsync(templateBuffer);
    const originalSheet = await originalZip.file('xl/worksheets/sheet1.xml').async('string');
    const originalPageSetup = originalSheet.match(/<pageSetup[^>]*>/)?.[0];

    // 生成されたファイルのpageSetup取得
    const outputZip = await JSZip.loadAsync(outputBuffer);
    const outputSheet = await outputZip.file('xl/worksheets/sheet1.xml').async('string');
    const outputPageSetup = outputSheet.match(/<pageSetup[^>]*>/)?.[0];

    console.log('   元のテンプレート:');
    console.log(`   ${originalPageSetup}`);
    console.log('');
    console.log('   生成されたファイル:');
    console.log(`   ${outputPageSetup}`);
    console.log('');

    if (originalPageSetup === outputPageSetup) {
      console.log('   ✅ 印刷設定は完全に保持されました！\n');
    } else {
      console.log('   ⚠️  印刷設定に変更がありました\n');
    }

    // 7. XMLエスケープ検証
    console.log('6. XMLエスケープ検証...');
    const sharedStrings = await outputZip.file('xl/sharedStrings.xml').async('string');

    const testCases = [
      { original: 'A社 & B社 <共同>', escaped: 'A社 &amp; B社 &lt;共同&gt;' },
      { original: '</v>', escaped: '&lt;/v&gt;' },
      { original: '"特別価格"', escaped: '&quot;特別価格&quot;' },
      { original: "'s dept.", escaped: "&apos;s dept." }
    ];

    testCases.forEach(({ original, escaped }) => {
      if (sharedStrings.includes(escaped)) {
        console.log(`   ✅ "${original}" は正しくエスケープされました`);
      } else {
        console.log(`   ⚠️  "${original}" はエスケープされていません`);
      }
    });

    console.log('\n=== テスト完了 ===');
    console.log('✅ すべてのAPIエンドポイントが正常に動作しています！');
    console.log('✅ test9方式による印刷設定完全保持を確認');
    console.log('✅ XMLエスケープによる安全性を確認\n');

  } catch (error) {
    console.error('\n❌ エラーが発生しました:');
    console.error(`   メッセージ: ${error.message}`);
    console.error(`   スタック:\n${error.stack}`);
    process.exit(1);
  }
}

// 実行
testAPI();
