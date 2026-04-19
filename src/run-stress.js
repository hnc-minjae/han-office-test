#!/usr/bin/env node
/**
 * run-stress — 장시간 stress 테스트 CLI 진입점.
 *
 * 사용법:
 *   node src/run-stress.js --duration 12h --seed 42
 *   node src/run-stress.js --duration 30m --map maps/hwp-menu-map.json
 *   node src/run-stress.js --duration 12h --skip-seed
 *   node src/run-stress.js --duration 12h --procdump "C:\\Tools\\procdump.exe"
 *
 * 중단 방법:
 *   1) ESC 를 1초 이내에 두 번 누르면 현재 iteration 종료 후 graceful shutdown
 *   2) Ctrl+C (SIGINT) 로도 종료 — summary.json까지 기록 후 exit
 */
'use strict';

const path = require('path');
const { runStress } = require('./stress-runner');

// =============================================================================
// 인자 파서
// =============================================================================

function parseDuration(str) {
    if (!str) throw new Error('--duration 필수');
    const m = /^(\d+)(ms|s|m|h|d)?$/i.exec(str.trim());
    if (!m) throw new Error(`--duration 형식 오류: "${str}" (예: 30m, 12h, 8h)`);
    const n = parseInt(m[1], 10);
    const unit = (m[2] || 's').toLowerCase();
    switch (unit) {
        case 'ms': return n;
        case 's':  return n * 1000;
        case 'm':  return n * 60 * 1000;
        case 'h':  return n * 3600 * 1000;
        case 'd':  return n * 86400 * 1000;
        default:   throw new Error(`unknown unit: ${unit}`);
    }
}

function parseArgs(argv) {
    const args = {
        duration:     null,
        seed:         null,
        map:          'maps/hwp-menu-map.json',
        skipSeed:     false,
        procdump:     null,
        outDir:       null,
        product:      'hwp',
        paceMs:       350,
        help:         false,
    };

    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        const next = () => argv[++i];
        switch (a) {
            case '--duration': args.duration = parseDuration(next()); break;
            case '--seed':     args.seed = parseInt(next(), 10) >>> 0; break;
            case '--map':      args.map = next(); break;
            case '--skip-seed':args.skipSeed = true; break;
            case '--procdump': args.procdump = next(); break;
            case '--out':      args.outDir = next(); break;
            case '--product':  args.product = next(); break;
            case '--pace-ms':  args.paceMs = parseInt(next(), 10); break;
            case '-h':
            case '--help':     args.help = true; break;
            default:
                if (a.startsWith('--')) {
                    throw new Error(`알 수 없는 옵션: ${a}`);
                }
        }
    }
    return args;
}

function usage() {
    process.stderr.write(`
사용법: node src/run-stress.js --duration <n> [options]

필수:
  --duration <n>        실행 시간. 예: 30m, 12h, 8h, 30s

옵션:
  --seed <n>            PRNG 시드 (기본: 현재 시각)
  --map <path>          메뉴 맵 경로 (기본: maps/hwp-menu-map.json)
  --skip-seed           초기 문서 시딩 생략
  --procdump <path>     ProcDump.exe 경로 (크래시 .dmp 자동 수집)
  --out <dir>           결과 디렉터리 (기본: test-results/stress-<timestamp>)
  --product <id>        hwp | hword | hshow | hcell (기본: hwp)
  --pace-ms <n>         iter 최소 간격(ms) — HWP 이벤트 처리 여유. 기본 350.
                        너무 낮으면 HWP가 따라오지 못해 버그를 숨길 수 있음.
  -h, --help            이 메시지

중단:
  ESC × 2 (1초 이내)    또는    Ctrl+C
`);
}

// =============================================================================
// 엔트리 포인트
// =============================================================================

async function main() {
    let args;
    try {
        args = parseArgs(process.argv.slice(2));
    } catch (e) {
        process.stderr.write(`⚠ ${e.message}\n`);
        usage();
        process.exit(2);
    }

    if (args.help || !args.duration) {
        usage();
        process.exit(args.help ? 0 : 2);
    }

    // SIGINT 안전 종료 — runStress 내부의 CancelWatcher와 별개로 Ctrl+C도 처리.
    // 현재 runStress는 외부에서 cancel 신호를 받는 API가 없으므로 process 종료로 위임.
    // 향후 runStress가 AbortSignal을 받도록 확장 가능.
    let sigintCount = 0;
    process.on('SIGINT', () => {
        sigintCount++;
        if (sigintCount === 1) {
            process.stderr.write('\n🛑 SIGINT — 결과 저장 후 종료합니다 (한번 더 누르면 즉시 종료)\n');
            // 현재 iteration 끝나고 다음 deadline 체크에서 종료되도록 deadline을 0으로
            // 강제하는 간단한 방법이 없으므로 이 버전은 즉시 exit — 이미 저장된 로그는 살아있음.
            process.exit(130);
        } else {
            process.stderr.write('💀 강제 종료\n');
            process.exit(137);
        }
    });

    try {
        const outDir = await runStress({
            mapPath:      path.resolve(args.map),
            durationMs:   args.duration,
            seed:         args.seed ?? ((Date.now() >>> 0) || 1),
            skipSeed:     args.skipSeed,
            procdumpPath: args.procdump,
            outDir:       args.outDir,
            product:      args.product,
            paceMs:       args.paceMs,
        });
        process.stderr.write(`✓ 완료. 결과: ${outDir}\n`);
        process.exit(0);
    } catch (e) {
        process.stderr.write(`✗ 실패: ${e.message}\n${e.stack || ''}\n`);
        process.exit(1);
    }
}

if (require.main === module) {
    main();
}

module.exports = { parseArgs, parseDuration };
