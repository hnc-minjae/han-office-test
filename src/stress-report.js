#!/usr/bin/env node
/**
 * stress-report — runStress 결과 디렉터리에서 actions.ndjson을 집계하여
 *                 report.md 를 생성한다.
 *
 * 사용법:
 *   node src/stress-report.js <outDir>
 *
 * 출력:
 *   <outDir>/report.md   요약 리포트
 *
 * 목적:
 *   - 전체 pass/fail/unknown 비율
 *   - Tab/category 별 비율 — 어느 영역이 약한가
 *   - 최하위 케이스 Top-20 — 어떤 맵 엔트리가 반복적으로 실패하는가
 *   - fail signals 빈도 — 어떤 패턴으로 깨지는가 (modified/pageCount/fields…)
 *   - 에러 발생 케이스 — 런타임 예외가 난 곳
 *
 * 이 리포트를 보면서 map-to-cases, menu-mapper 프로빙, verifier 분류 규칙을
 * 반복 개선한다 (V2, V3…).
 */
'use strict';

const fs = require('fs');
const path = require('path');

function readNdjson(filePath) {
    if (!fs.existsSync(filePath)) return [];
    const raw = fs.readFileSync(filePath, 'utf8');
    const out = [];
    for (const line of raw.split(/\r?\n/)) {
        if (!line.trim()) continue;
        try { out.push(JSON.parse(line)); } catch (_) {}
    }
    return out;
}

function pad(s, n) {
    s = String(s);
    if (s.length >= n) return s;
    return s + ' '.repeat(n - s.length);
}

function pct(num, den) {
    if (!den) return '-';
    return (100 * num / den).toFixed(1) + '%';
}

// =============================================================================
// 집계
// =============================================================================

function aggregate(entries) {
    const summary = {
        totalIter: entries.length,
        byKind: {},
        ribbon: {
            total: 0, pass: 0, fail: 0, unknown: 0, error: 0,
            byTab: {},           // tab -> { pass, fail, unknown }
            byCategory: {},      // category -> { pass, fail, unknown }
            byCase: {},          // caseId -> { tab, name, category, total, pass, fail, unknown, reasons: Set }
            byRule: {},          // V2 action rule -> { pass, fail, unknown, total }
            failSignals: {},     // signal -> count
            failSamples: [],     // {iter, caseId, signals, reason}[]
            errorSamples: [],    // {iter, caseId, error}[]
        },
    };

    for (const e of entries) {
        summary.byKind[e.actionType] = (summary.byKind[e.actionType] || 0) + 1;

        if (e.actionType !== 'ribbon' || !e.caseId) continue;
        const r = summary.ribbon;
        r.total++;

        if (e.error) {
            r.error++;
            if (r.errorSamples.length < 30) {
                r.errorSamples.push({ iter: e.iter, caseId: e.caseId, error: e.error });
            }
            continue;
        }

        const ver = e.verification;
        if (!ver) continue;
        const verdict = ver.verdict?.verdict || 'unknown';
        r[verdict] = (r[verdict] || 0) + 1;

        const tab = e.tab || '(none)';
        r.byTab[tab] = r.byTab[tab] || { pass: 0, fail: 0, unknown: 0 };
        r.byTab[tab][verdict] = (r.byTab[tab][verdict] || 0) + 1;

        const cat = e.category || '(none)';
        r.byCategory[cat] = r.byCategory[cat] || { pass: 0, fail: 0, unknown: 0 };
        r.byCategory[cat][verdict] = (r.byCategory[cat][verdict] || 0) + 1;

        const caseId = e.caseId;
        if (!r.byCase[caseId]) {
            r.byCase[caseId] = {
                tab: e.tab, name: e.name, shortName: e.name, category: e.category,
                total: 0, pass: 0, fail: 0, unknown: 0,
                reasons: {},  // reason -> count
            };
        }
        const entry = r.byCase[caseId];
        entry.total++;
        entry[verdict]++;
        if (ver.verdict?.reason) {
            entry.reasons[ver.verdict.reason] = (entry.reasons[ver.verdict.reason] || 0) + 1;
        }

        if (verdict === 'fail') {
            const signals = ver.verdict?.signals || [];
            for (const s of signals) {
                r.failSignals[s] = (r.failSignals[s] || 0) + 1;
            }
            if (r.failSamples.length < 30) {
                r.failSamples.push({
                    iter: e.iter, caseId,
                    signals, reason: ver.verdict?.reason || '',
                    interaction: e.interaction,
                });
            }
        }

        // V2: action rule 집계
        const rule = ver.verdict?.rule;
        if (rule) {
            r.byRule[rule] = r.byRule[rule] || { pass: 0, fail: 0, unknown: 0, total: 0 };
            r.byRule[rule].total++;
            r.byRule[rule][verdict] = (r.byRule[rule][verdict] || 0) + 1;
        }
    }
    return summary;
}

// =============================================================================
// 리포트 생성
// =============================================================================

function renderReport(summary, outDir, configPath, summaryJsonPath) {
    const lines = [];
    const h = (s) => { lines.push(s); };

    h('# Stress Test Report');
    h('');
    h(`- outDir: \`${outDir}\``);
    if (fs.existsSync(configPath)) {
        try {
            const cfg = JSON.parse(fs.readFileSync(configPath, 'utf8'));
            h(`- runId: \`${cfg.runId}\`  seed: \`${cfg.seed}\`  duration: ${cfg.durationMs}ms`);
            h(`- startedAt: ${cfg.startedAt}`);
        } catch (_) {}
    }
    if (fs.existsSync(summaryJsonPath)) {
        try {
            const js = JSON.parse(fs.readFileSync(summaryJsonPath, 'utf8'));
            h(`- iterations: ${js.iterations}  crashes: ${js.crashes}  errors: ${js.errors}  elapsedMs: ${js.elapsedMs}`);
            if (js.reason) h(`- endReason: ${js.reason}`);
        } catch (_) {}
    }
    h('');

    // byKind
    h('## Iteration 타입 분포');
    h('');
    h('| type | count |');
    h('|---|---|');
    for (const [k, v] of Object.entries(summary.byKind).sort((a, b) => b[1] - a[1])) {
        h(`| ${k} | ${v} |`);
    }
    h('');

    const r = summary.ribbon;

    // ribbon 전체
    h('## Ribbon 액션 검증 결과');
    h('');
    h(`- 총 ${r.total}회 (pass ${r.pass} / fail ${r.fail} / unknown ${r.unknown} / error ${r.error})`);
    h(`- pass율: ${pct(r.pass, r.total)}   fail율: ${pct(r.fail, r.total)}   unknown율: ${pct(r.unknown, r.total)}`);
    h('');

    // tab별
    h('### Tab별 verdict');
    h('');
    h('| tab | pass | fail | unknown | fail율 |');
    h('|---|---:|---:|---:|---:|');
    const tabs = Object.entries(r.byTab).sort((a, b) => (b[1].fail / (b[1].pass+b[1].fail+b[1].unknown||1)) - (a[1].fail / (a[1].pass+a[1].fail+a[1].unknown||1)));
    for (const [tab, stat] of tabs) {
        const total = stat.pass + stat.fail + stat.unknown;
        h(`| ${tab} | ${stat.pass} | ${stat.fail} | ${stat.unknown} | ${pct(stat.fail, total)} |`);
    }
    h('');

    // category별
    h('### Category별 verdict');
    h('');
    h('| category | pass | fail | unknown | fail율 |');
    h('|---|---:|---:|---:|---:|');
    for (const [cat, stat] of Object.entries(r.byCategory)) {
        const total = stat.pass + stat.fail + stat.unknown;
        h(`| ${cat} | ${stat.pass} | ${stat.fail} | ${stat.unknown} | ${pct(stat.fail, total)} |`);
    }
    h('');

    // fail signal 빈도
    h('### Fail signal 빈도');
    h('');
    h('fail verdict의 원인으로 가장 자주 등장한 시그널 — V2에서 어떤 패턴부터 정교화할지 결정하는 근거.');
    h('');
    h('| signal | count |');
    h('|---|---:|');
    for (const [s, c] of Object.entries(r.failSignals).sort((a, b) => b[1] - a[1])) {
        h(`| ${s} | ${c} |`);
    }
    h('');

    // 최하위 케이스 (fail 율 높음 + 최소 호출수)
    h('### 최하위 20개 케이스 (fail율 상위)');
    h('');
    const caseList = Object.entries(r.byCase)
        .map(([id, s]) => ({
            id, ...s,
            failRate: s.fail / (s.total || 1),
        }))
        .filter(c => c.total >= 3)
        .sort((a, b) => b.failRate - a.failRate)
        .slice(0, 20);
    if (caseList.length === 0) {
        h('(충분한 호출 수를 가진 fail 케이스 없음)');
    } else {
        h('| case | tab | category | total | pass | fail | fail율 | top reason |');
        h('|---|---|---|---:|---:|---:|---:|---|');
        for (const c of caseList) {
            const topReason = Object.entries(c.reasons).sort((a, b) => b[1] - a[1])[0];
            const reasonStr = topReason ? `${topReason[0]} ×${topReason[1]}` : '-';
            h(`| \`${c.id}\` | ${c.tab} | ${c.category} | ${c.total} | ${c.pass} | ${c.fail} | ${pct(c.fail, c.total)} | ${reasonStr} |`);
        }
    }
    h('');

    // unknown 케이스 Top (V2 분류 규칙 추가 우선순위)
    h('### Unknown 비율 높은 케이스 Top 20');
    h('');
    h('verifier V1이 분류하지 못한 케이스. V2에서 분류 규칙을 추가할 우선순위.');
    h('');
    const unknownList = Object.entries(r.byCase)
        .map(([id, s]) => ({ id, ...s, unkRate: s.unknown / (s.total || 1) }))
        .filter(c => c.total >= 3 && c.unknown > 0)
        .sort((a, b) => b.unkRate - a.unkRate || b.total - a.total)
        .slice(0, 20);
    if (unknownList.length === 0) {
        h('(unknown 케이스 없음)');
    } else {
        h('| case | tab | category | total | unknown | unknown율 |');
        h('|---|---|---|---:|---:|---:|');
        for (const c of unknownList) {
            h(`| \`${c.id}\` | ${c.tab} | ${c.category} | ${c.total} | ${c.unknown} | ${pct(c.unknown, c.total)} |`);
        }
    }
    h('');

    // 에러 샘플
    h('### 런타임 에러 샘플');
    h('');
    if (r.errorSamples.length === 0) {
        h('(없음)');
    } else {
        h('| iter | caseId | error |');
        h('|---:|---|---|');
        for (const s of r.errorSamples.slice(0, 20)) {
            h(`| ${s.iter} | \`${s.caseId}\` | ${s.error.replace(/\|/g, '\\|')} |`);
        }
    }
    h('');

    // fail 샘플
    h('### Fail 샘플 (첫 30개)');
    h('');
    if (r.failSamples.length === 0) {
        h('(없음)');
    } else {
        h('| iter | caseId | interaction | signals | reason |');
        h('|---:|---|---|---|---|');
        for (const s of r.failSamples) {
            h(`| ${s.iter} | \`${s.caseId}\` | ${s.interaction || '-'} | ${s.signals.join(', ')} | ${s.reason.replace(/\|/g, '\\|')} |`);
        }
    }
    h('');

    // 규칙(rule) 히트 분포 — V2 action 분류 상태 확인
    if (Object.keys(r.byRule).length > 0) {
        h('### Action Rule 히트 분포 (V2)');
        h('');
        h('Action 카테고리에서 매칭된 규칙별 분류 결과.');
        h('');
        h('| rule | total | pass | fail | unknown | pass율 |');
        h('|---|---:|---:|---:|---:|---:|');
        for (const [rule, s] of Object.entries(r.byRule).sort((a, b) => b[1].total - a[1].total)) {
            h(`| \`${rule}\` | ${s.total} | ${s.pass} | ${s.fail} | ${s.unknown} | ${pct(s.pass, s.total)} |`);
        }
        h('');
    }

    // 다음 개선 제안 (자동 생성 힌트)
    h('## V2 개선 제안');
    h('');
    if (r.failSignals['modified']) {
        h(`- **"dialog cancel → modified" 가 ${r.failSignals['modified']}건**: cleanup Esc가 다이얼로그의 변경을 완전히 취소하지 못하는 케이스. execRibbonCase의 teardown을 수정(Esc 2회 + 명시적 '취소' 버튼 클릭)하면 감소할 것.`);
    }
    const highUnknown = (r.unknown / r.total) > 0.2;
    if (highUnknown) {
        h(`- **unknown 비율 ${pct(r.unknown, r.total)}**: action 카테고리 분류 규칙이 없음. 이름 기반 규칙(예: "붙이기" → 텍스트 늘어남 기대)을 verifier에 추가하면 정확도 상승.`);
    }
    if (r.errorSamples.length > 0) {
        h(`- **런타임 에러 ${r.error}건**: ${r.errorSamples.length}개 샘플 조사하여 원인 수정.`);
    }
    h('');

    return lines.join('\n');
}

// =============================================================================
// CLI
// =============================================================================

function main() {
    const outDir = process.argv[2];
    if (!outDir) {
        process.stderr.write('사용법: node src/stress-report.js <outDir>\n');
        process.exit(2);
    }
    if (!fs.existsSync(outDir)) {
        process.stderr.write(`경로 없음: ${outDir}\n`);
        process.exit(2);
    }

    const actionsPath = path.join(outDir, 'actions.ndjson');
    const entries = readNdjson(actionsPath);
    if (entries.length === 0) {
        process.stderr.write(`actions.ndjson 비어있음: ${actionsPath}\n`);
        process.exit(2);
    }

    const summary = aggregate(entries);
    const reportPath = path.join(outDir, 'report.md');
    const md = renderReport(summary, outDir,
        path.join(outDir, 'config.json'),
        path.join(outDir, 'summary.json'));
    fs.writeFileSync(reportPath, md, 'utf8');
    process.stderr.write(`✓ 리포트 생성: ${reportPath}\n`);
    process.stderr.write(`  총 iter: ${summary.totalIter}  ribbon: ${summary.ribbon.total}\n`);
    process.stderr.write(`  ribbon verdict — pass: ${summary.ribbon.pass}  fail: ${summary.ribbon.fail}  unknown: ${summary.ribbon.unknown}  error: ${summary.ribbon.error}\n`);
}

if (require.main === module) {
    main();
}

module.exports = { aggregate, renderReport, readNdjson };
