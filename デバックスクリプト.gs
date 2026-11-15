/**
 * ============================================================================
 * 🔍 拠点名・住所 100%特定検証デバッグスクリプト（完全非破壊・読取専用）
 * 1. 拠点マスタ全件の「住所登録有無」＆「M3形式（番地カット町名）での自己逆引きテスト」
 * 2. 採用くん（処理済み・紹介会社進捗）全履歴の拠点・住所同定テスト
 * 3. シフト系4シート（貼付用・確定・応募・募集）全行の拠点同定テスト
 * 4. saiyokun側のM3住所変換ロジックとの一致比較チェック
 * ============================================================================
 */
function debugVerifyAllClinicIdentifications() {
  const startTime = Date.now();
  Logger.log("=== 🔍 拠点名・住所 100%特定検証デバッグ 開始 ===");

  if (typeof clearCheckerRuntimeCache === "function") {
    clearCheckerRuntimeCache();
  }

  const PASTE_MASTER_ID   = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA';
  const SHIFT_MASTER_ID   = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const LOC_MASTER_ID     = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const RECRUIT_MASTER_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ';

  const locSs     = getCachedSs(LOC_MASTER_ID);
  const recruitSs = getCachedSs(RECRUIT_MASTER_ID);
  const pasteSs   = getCachedSs(PASTE_MASTER_ID);
  const shiftSs   = getCachedSs(SHIFT_MASTER_ID);

  // 業務枠（クリニックではない特殊枠）の除外定義
  const EXCLUDED_NON_CLINICS = [
    '院外勤務', '嘱託医業務', '医師会業務', 'バックアップ', '有給', '有休', '欠勤', '休館'
  ];
  const isExcludedTask = (raw) => {
    const s = String(raw || '').trim();
    return EXCLUDED_NON_CLINICS.some(ex => s.includes(ex));
  };

  // Phase4-A の本番エンジンをそのまま構築
  const alertkunEngine = buildPhase4ClinicIdentifier_(locSs, EXCLUDED_NON_CLINICS);

  // ------------------------------------------------------------------
  // 比較用：saiyokun (1-B1) 側のエムスリー住所変換ロジックを再現
  // ------------------------------------------------------------------
  const extractTownM3Style = (rawAddr) => {
    if (!rawAddr) return '';
    let s = String(rawAddr)
      .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/[\s ]+/g, '');
    const m = s.match(/^([^\d\-ー−―‐ｰ]+)/);
    let town = m ? m[1] : s;
    const chomeMatch = s.match(/^([^\d\-ー−―‐ｰ]+?[一二三四五六七八九十]+丁目)/);
    if (chomeMatch) {
      town = chomeMatch[1];
    } else {
      town = town.replace(/[一二三四五六七八九十]+$/, '');
    }
    return town;
  };

  // ==================================================================
  // 【TEST 1】拠点マスタ自体の健全性 ＆ M3形式（町名切り出し）逆引きテスト
  // ==================================================================
  Logger.log("\n------------------------------------------------------------");
  Logger.log("📋 【TEST 1】拠点マスタ全行の住所登録 ＆ M3住所逆引きテスト");
  Logger.log("------------------------------------------------------------");

  const locSheet = locSs.getSheetByName("拠点名");
  const locData = getCachedDisplayValues(locSheet);
  const locHeaders = locData[0].map(h => String(h).replace(/[\s ]+/g, ''));
  const idxFormal = locHeaders.indexOf("正規記載");
  const idxStatus = locHeaders.indexOf("システム反映");
  const idxAddr   = locHeaders.findIndex(h => h.includes("住所"));

  const saiyokunMap = [];
  let masterCount = 0;
  const missingAddrClinics = [];
  const reverseMismatchClinics = [];

  for (let i = 1; i < locData.length; i++) {
    const formal = String(locData[i][idxFormal] || "").replace(/[\s ]+/g, "");
    if (!formal || isExcludedTask(formal)) continue;
    masterCount++;

    const sysStatus = idxStatus !== -1 ? String(locData[i][idxStatus]).trim() : "";
    const rawAddr   = idxAddr !== -1 ? String(locData[i][idxAddr]).trim() : "";

    const townName = extractTownM3Style(rawAddr);
    const townWithoutChome = townName.replace(/[一二三四五六七八九十]+丁目$/, '');
    saiyokunMap.push({
      clinicName: formal.replace(/キャップスクリニック/g, ''),
      rawAddr,
      townName,
      townWithoutChome
    });

    if (!rawAddr) {
      missingAddrClinics.push(`${formal} (システム反映:${sysStatus || '未設定'})`);
      continue;
    }

    // ① フル住所での逆引きチェック
    const resFull = alertkunEngine.identifyClinic(rawAddr);
    // ② エムスリー応募メールと同じ「番地以降をカットした住所＋（小児科）」形式での逆引きチェック
    const m3SimulatedInput = rawAddr.replace(/\d.*$/, '').trim() + '（小児科）';
    const resM3Sim = alertkunEngine.identifyClinic(m3SimulatedInput);

    if (!resFull.resolved || resFull.formal !== formal || !resM3Sim.resolved || resM3Sim.formal !== formal) {
      reverseMismatchClinics.push(
        `  ⚠️ 拠点[${formal}] 住所[${rawAddr}]\n` +
        `     ・フル住所判定 -> ${resFull.formal} (resolved:${resFull.resolved})\n` +
        `     ・M3模擬[${m3SimulatedInput}] -> ${resM3Sim.formal} (resolved:${resM3Sim.resolved})`
      );
    }
  }

  // saiyokun側の変換関数（比較用）
  const convertBySaiyokunM3 = (rawAddr) => {
    const addr = String(rawAddr || '').trim();
    const cleanAddr = addr.replace(/[（(][^）)]*(内科|小児科)[^）)]*[）)]/g, '').trim();
    const inputTown = extractTownM3Style(cleanAddr);
    const inputTownNoChome = inputTown.replace(/[一二三四五六七八九十]+丁目$/, '');

    const sortedByTown = [...saiyokunMap].sort((a, b) => b.townName.length - a.townName.length);
    for (const item of sortedByTown) {
      if (!item.townName || item.townName.length < 3) continue;
      if (inputTown.includes(item.townName) || item.townName.includes(inputTown)) return item.clinicName;
    }
    const sortedByNoChome = [...saiyokunMap].sort((a, b) => b.townWithoutChome.length - a.townWithoutChome.length);
    for (const item of sortedByNoChome) {
      if (!item.townWithoutChome || item.townWithoutChome.length < 3) continue;
      if (inputTownNoChome.includes(item.townWithoutChome) || item.townWithoutChome.includes(inputTownNoChome)) return item.clinicName;
    }
    const sortedByName = [...saiyokunMap].sort((a, b) => b.clinicName.length - a.clinicName.length);
    for (const item of sortedByName) {
      if (item.clinicName && item.clinicName.length >= 2 && cleanAddr.includes(item.clinicName)) return item.clinicName;
    }
    return null;
  };

  Logger.log(`✅ マスタ登録拠点数: ${masterCount} 拠点`);
  if (missingAddrClinics.length === 0) {
    Logger.log(`✅ 住所空欄の拠点: 0 件（全拠点に住所が登録されています）`);
  } else {
    Logger.log(`⚠️ 住所空欄の拠点 (${missingAddrClinics.length}件): ${missingAddrClinics.join(", ")}`);
  }

  if (reverseMismatchClinics.length === 0) {
    Logger.log(`✅ フル住所＆M3町名形式からの自己逆引き: 全拠点 100% 一致！`);
  } else {
    Logger.log(`🚨 住所逆引きで不一致または曖昧になった拠点 (${reverseMismatchClinics.length}件):`);
    reverseMismatchClinics.forEach(msg => Logger.log(msg));
  }

  // ==================================================================
  // 共通スキャン関数（ユニークな表記をすべて集めて特定テスト）
  // ==================================================================
  const inspectSheetClinicColumn = (label, sheet, targetColKeywords) => {
    if (!sheet) {
      Logger.log(`⚠️ [${label}] シートが見つかりません`);
      return;
    }
    const data = getCachedDisplayValues(sheet);
    if (data.length < 2) {
      Logger.log(`ℹ️ [${label}] データ行なし`);
      return;
    }

    const headers = data[0].map(h => String(h).replace(/\r?\n/g, '').trim());
    let colIdx = -1;
    for (const kw of targetColKeywords) {
      colIdx = headers.findIndex(h => h === kw || h.includes(kw));
      if (colIdx !== -1) break;
    }
    if (colIdx === -1) {
      Logger.log(`⚠️ [${label}] 拠点列が見つかりません`);
      return;
    }

    const uniqueRawCounts = new Map(); // rawStr -> { count, sampleRow }
    let totalValidRows = 0;

    for (let r = 1; r < data.length; r++) {
      const rawVal = String(data[r][colIdx] || "").trim();
      if (!rawVal || isExcludedTask(rawVal)) continue;
      totalValidRows++;
      if (!uniqueRawCounts.has(rawVal)) {
        uniqueRawCounts.set(rawVal, { count: 1, sampleRow: r + 1 });
      } else {
        uniqueRawCounts.get(rawVal).count++;
      }
    }

    const resolvedSamples = [];
    const unresolvedList = [];
    const engineDiffList = [];

    for (const [rawVal, info] of uniqueRawCounts.entries()) {
      const res = alertkunEngine.identifyClinic(rawVal);
      if (!res.resolved) {
        unresolvedList.push(`  ❌ 未特定: "${rawVal}" (${info.count}行, 例:行${info.sampleRow})`);
      } else {
        // 元の文字列と正規記載が異なる「住所表記・揺れ表記」の変換結果を一覧用に収集
        if (rawVal.replace(/[\s ]+/g, '') !== res.formal) {
          resolvedSamples.push({ raw: rawVal, formal: res.formal, count: info.count });
        }

        // もし都道府県や市区郡を含む「住所らしき文字列」なら、saiyokun側の変換結果とも突き合わせる
        if (/[都道府県市区町村]/.test(rawVal) && !rawVal.includes('千葉ニュータウン')) {
          const saiyoConverted = convertBySaiyokunM3(rawVal);
          if (!saiyoConverted || saiyoConverted !== res.formal) {
            engineDiffList.push(`  ⚡ エンジン差: "${rawVal}" -> Alertkun:[${res.formal}] / saiyokun(1-B1):[${saiyoConverted || '変換不可'}]`);
          }
        }
      }
    }

    Logger.log(`\n------------------------------------------------------------`);
    Logger.log(`📂 【${label}】 対象行数: ${totalValidRows} 行 （ユニーク表記: ${uniqueRawCounts.size} パターン）`);
    Logger.log(`------------------------------------------------------------`);

    if (unresolvedList.length === 0) {
      Logger.log(`🎉 特定率 100%！（全 ${uniqueRawCounts.size} パターンすべて正規拠点に同定成功）`);
    } else {
      Logger.log(`🚨 未特定の表記が ${unresolvedList.length} パターンありました:`);
      unresolvedList.forEach(line => Logger.log(line));
    }

    if (resolvedSamples.length > 0) {
      resolvedSamples.sort((a, b) => b.count - a.count);
      const showLimit = Math.min(resolvedSamples.length, 30);
      Logger.log(`🔎 表記揺れ・住所からの変換実例（全${resolvedSamples.length}パターン中 上位${showLimit}件）:`);
      resolvedSamples.slice(0, showLimit).forEach(s => {
        Logger.log(`   ・"${s.raw}"  ==>  【${s.formal}】 (${s.count}件)`);
      });
    }

    if (engineDiffList.length > 0) {
      Logger.log(`💡 参考：saiyokun(1-B1)の簡易町名変換と差が出た住所表記 (${engineDiffList.length}件):`);
      engineDiffList.forEach(d => Logger.log(d));
    }
  };

  // ==================================================================
  // 【TEST 2】採用くん（処理済み・紹介会社進捗）全履歴スキャン
  // ==================================================================
  let saiyoArchiveSheet = null;
  recruitSs.getSheets().forEach(s => {
    if (s.getSheetId() == '1155146432' || s.getName().includes("処理済み")) saiyoArchiveSheet = s;
  });
  inspectSheetClinicColumn("採用くん：処理済みシート（全履歴）", saiyoArchiveSheet, ["拠点名", "拠点"]);
  inspectSheetClinicColumn("採用くん：紹介会社進捗シート", recruitSs.getSheetByName("紹介会社進捗"), ["拠点名", "拠点"]);

  // ==================================================================
  // 【TEST 3】シフト系4シート（貼付用・確定・応募・募集）全行スキャン
  // ==================================================================
  inspectSheetClinicColumn("貼付マスタ：貼付用（実績）", pasteSs.getSheetByName("貼付用"), ["クリニック名", "拠点名"]);
  inspectSheetClinicColumn("シフトマスタ：確定シフト", shiftSs.getSheetByName("確定シフト"), ["クリニック名", "拠点名"]);
  inspectSheetClinicColumn("シフトマスタ：応募シフト", shiftSs.getSheetByName("応募シフト"), ["クリニック名", "拠点名"]);
  inspectSheetClinicColumn("シフトマスタ：募集シフト", shiftSs.getSheetByName("募集シフト"), ["クリニック名", "拠点名"]);

  const sec = ((Date.now() - startTime) / 1000).toFixed(2);
  Logger.log(`\n=== 🏁 拠点名・住所 特定検証完了（所要時間: ${sec}秒） ===`);
}