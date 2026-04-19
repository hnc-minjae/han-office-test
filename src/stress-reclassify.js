#!/usr/bin/env node
/**
 * stress-reclassify — 기존 actions.ndjson의 각 entry를 현재 verifier.classify로
 * 재분류하여 새 디렉터리에 actions.ndjson + report.md 를 생성.
 *
 * 용도: verifier 규칙이 업데이트됐을 때 이전 run의 raw 데이터를 활용해
 *       즉시 새 classification 결과를 얻는다. 30분 재실행 비용 0.
 *
 * 제약: entry.verification.{diff, before, after}에 의존. V1 이전 run은 지원 안 됨.
 *
 * 사용법:
 *   node src/stress-reclassify.js <inputDir> [outputDir]
 *   # outputDir 생략 시 <inputDir>-reclassified
 */
'use strict';

const fs = require('fs');
const path = require('path');
const verifier = require('./stress-verifier');
const { aggregate, renderReport } = require('./stress-report');

function readNdjson(fp) {
    if (!fs.existsSync(fp)) return [];
    return fs.readFileSync(fp, 'utf8').split(/\r?\n/)
        .filter(l => l.trim())
        .map(l => { try { return JSON.parse(l); } catch (_) { return null; } })
        .filter(Boolean);
}

function main() {
    const inputDir = process.argv[2];
    if (!inputDir) {
        process.stderr.write('사용법: node src/stress-reclassify.js <inputDir> [outputDir]\n');
        process.exit(2);
    }
    if (!fs.existsSync(inputDir)) {
        process.stderr.write(`입력 경로 없음: ${inputDir}\n`);
        process.exit(2);
    }
    const outputDir = process.argv[3] || `${inputDir}-reclassified`;
    fs.mkdirSync(outputDir, { recursive: true });

    const inputActionsPath = path.join(inputDir, 'actions.ndjson');
    const outputActionsPath = path.join(outputDir, 'actions.ndjson');

    const entries = readNdjson(inputActionsPath);
    if (!entries.length) {
        process.stderr.write(`actions.ndjson 비어있음: ${inputActionsPath}\n`);
        process.exit(2);
    }

    process.stderr.write(`▶ 재분류: ${entries.length}개 entry\n`);

    const out = [];
    const stats = { reclassified: 0, skipped: 0 };
    for (const e of entries) {
        const ver = e.verification;
        if (!ver || !ver.diff || e.actionType !== 'ribbon') {
            out.push(e);
            stats.skipped++;
            continue;
        }
        const caseObj = {
            id: e.caseId,
            tab: e.tab,
            name: e.name,
            shortName: e.name,
            category: e.category,
        };
        const newVerdict = verifier.classify(caseObj, e.interaction || 'none', ver.diff, ver.before, ver.after);
        const updated = { ...e, verification: { ...ver, verdict: newVerdict } };
        out.push(updated);
        stats.reclassified++;
    }

    fs.writeFileSync(outputActionsPath, out.map(x => JSON.stringify(x)).join('\n') + '\n');

    // config.json, summary.json 복사 (리포터가 참조)
    for (const fname of ['config.json', 'summary.json']) {
        const src = path.join(inputDir, fname);
        if (fs.existsSync(src)) fs.copyFileSync(src, path.join(outputDir, fname));
    }

    process.stderr.write(`✓ 재분류 완료: ${stats.reclassified} entries, ${stats.skipped} skipped\n`);

    // 리포트 생성
    const summary = aggregate(out);
    const md = renderReport(summary, outputDir,
        path.join(outputDir, 'config.json'),
        path.join(outputDir, 'summary.json'));
    fs.writeFileSync(path.join(outputDir, 'report.md'), md, 'utf8');
    process.stderr.write(`✓ 리포트: ${path.join(outputDir, 'report.md')}\n`);
    process.stderr.write(`  ribbon verdict — pass: ${summary.ribbon.pass}  fail: ${summary.ribbon.fail}  unknown: ${summary.ribbon.unknown}  error: ${summary.ribbon.error}\n`);
}

if (require.main === module) main();
