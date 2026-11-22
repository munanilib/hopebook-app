/* QR-BEGIN */
// 외부 서비스 없이 QR 코드를 만듭니다. (바이트 모드 · 오류정정 M · 버전 1~10, 최대 213바이트)
// 알고리즘은 QR 코드 표준(ISO/IEC 18004)을 따릅니다.
var QrCore = (function () {
  'use strict';

  // 버전별: [전체 코드워드, 블록 수, 블록당 오류정정 코드워드]  (오류정정 M)
  var TABLE = {
    1: [26, 1, 10], 2: [44, 1, 16], 3: [70, 1, 26], 4: [100, 2, 18], 5: [134, 2, 24],
    6: [172, 4, 16], 7: [196, 4, 18], 8: [242, 4, 22], 9: [292, 5, 22], 10: [346, 5, 26]
  };
  var MAX_VERSION = 10;

  // --- 갈루아체 GF(256) ---
  var EXP = [], LOG = [];
  (function () {
    var x = 1, i;
    for (i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 0x100) x ^= 0x11d; }
    for (i = 255; i < 512; i++) EXP[i] = EXP[i - 255];
  })();
  function gfMul(a, b) { return (a === 0 || b === 0) ? 0 : EXP[LOG[a] + LOG[b]]; }

  function rsGenerator(degree) {
    var poly = [1], i, j, next;
    for (i = 0; i < degree; i++) {
      next = [];
      for (j = 0; j <= poly.length; j++) next[j] = 0;
      for (j = 0; j < poly.length; j++) {
        next[j] ^= poly[j];
        next[j + 1] ^= gfMul(poly[j], EXP[i]);
      }
      poly = next;
    }
    return poly; // 최고차항 계수 1 포함, 길이 degree+1
  }
  function rsRemainder(data, gen) {
    var degree = gen.length - 1, rem = [], i, j, factor;
    for (i = 0; i < degree; i++) rem[i] = 0;
    for (i = 0; i < data.length; i++) {
      factor = data[i] ^ rem[0];
      rem.shift(); rem.push(0);
      for (j = 0; j < degree; j++) rem[j] ^= gfMul(gen[j + 1], factor);
    }
    return rem;
  }

  function utf8Bytes(text) {
    var out = [], i, c, d;
    var s = unescape(encodeURIComponent(String(text)));
    for (i = 0; i < s.length; i++) out.push(s.charCodeAt(i));
    return out;
  }

  function numAlign(ver) { return ver === 1 ? 0 : Math.floor(ver / 7) + 2; }
  function rawModules(ver) {
    var n = (16 * ver + 128) * ver + 64;
    if (ver >= 2) { var a = numAlign(ver); n -= (25 * a - 10) * a - 55; if (ver >= 7) n -= 36; }
    return n;
  }
  function alignPositions(ver) {
    if (ver === 1) return [];
    var n = numAlign(ver), size = ver * 4 + 17;
    var step = Math.ceil((ver * 4 + 4) / (n * 2 - 2)) * 2;
    var res = [6], pos;
    for (pos = size - 7; res.length < n; pos -= step) res.splice(1, 0, pos);
    return res;
  }

  function bit(v, i) { return ((v >>> i) & 1) !== 0; }

  function pickVersion(byteLen) {
    var ver, t, dataCw, capBits, need;
    for (ver = 1; ver <= MAX_VERSION; ver++) {
      t = TABLE[ver];
      dataCw = t[0] - t[1] * t[2];
      capBits = dataCw * 8;
      need = 4 + (ver <= 9 ? 8 : 16) + byteLen * 8;
      if (need <= capBits) return ver;
    }
    throw new Error('주소가 너무 길어서 QR을 만들 수 없습니다.');
  }

  function encode(text) {
    var bytes = utf8Bytes(text);
    var ver = pickVersion(bytes.length);
    var t = TABLE[ver], raw = Math.floor(rawModules(ver) / 8);
    var numBlocks = t[1], eccLen = t[2], dataCw = raw - numBlocks * eccLen;
    var size = ver * 4 + 17;
    var i, j, k;

    // 1) 데이터 비트열
    var bits = [];
    function push(val, len) { for (var b = len - 1; b >= 0; b--) bits.push((val >>> b) & 1); }
    push(4, 4);                               // 바이트 모드
    push(bytes.length, ver <= 9 ? 8 : 16);    // 글자 수
    for (i = 0; i < bytes.length; i++) push(bytes[i], 8);
    var capBits = dataCw * 8;
    for (i = 0; i < 4 && bits.length < capBits; i++) bits.push(0); // 종료 문자
    while (bits.length % 8 !== 0) bits.push(0);
    var data = [];
    for (i = 0; i < bits.length; i += 8) {
      var v = 0; for (j = 0; j < 8; j++) v = (v << 1) | bits[i + j]; data.push(v);
    }
    for (k = 0; data.length < dataCw; k++) data.push(k % 2 === 0 ? 0xEC : 0x11);

    // 2) 블록 나누기 + 오류정정 + 끼워 넣기
    var shortLen = Math.floor(raw / numBlocks), numShort = numBlocks - raw % numBlocks;
    var gen = rsGenerator(eccLen), blocks = [], offset = 0;
    for (i = 0; i < numBlocks; i++) {
      var len = shortLen - eccLen + (i < numShort ? 0 : 1);
      var dat = data.slice(offset, offset + len); offset += len;
      var ecc = rsRemainder(dat, gen);
      if (i < numShort) dat.push(0);          // 짧은 블록은 자리 맞춤용 빈칸
      blocks.push(dat.concat(ecc));
    }
    var all = [];
    for (i = 0; i < blocks[0].length; i++) {
      for (j = 0; j < blocks.length; j++) {
        if (i !== shortLen - eccLen || j >= numShort) all.push(blocks[j][i]);
      }
    }

    // 3) 모듈 배치
    var modules = [], isFn = [];
    for (i = 0; i < size; i++) { modules.push([]); isFn.push([]); for (j = 0; j < size; j++) { modules[i].push(false); isFn[i].push(false); } }
    function setFn(x, y, dark) { modules[y][x] = !!dark; isFn[y][x] = true; }

    for (i = 0; i < size; i++) { setFn(6, i, i % 2 === 0); setFn(i, 6, i % 2 === 0); }
    function finder(cx, cy) {
      for (var dy = -4; dy <= 4; dy++) for (var dx = -4; dx <= 4; dx++) {
        var dist = Math.max(Math.abs(dx), Math.abs(dy)), xx = cx + dx, yy = cy + dy;
        if (xx >= 0 && xx < size && yy >= 0 && yy < size) setFn(xx, yy, dist !== 2 && dist !== 4);
      }
    }
    finder(3, 3); finder(size - 4, 3); finder(3, size - 4);
    var ap = alignPositions(ver), n = ap.length;
    for (i = 0; i < n; i++) for (j = 0; j < n; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;
      for (var dy = -2; dy <= 2; dy++) for (var dx = -2; dx <= 2; dx++)
        setFn(ap[i] + dx, ap[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
    }
    function drawFormat(mask) {
      var d = (0 << 3) | mask;                // 오류정정 M = 00
      var rem = d;
      for (var q = 0; q < 10; q++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
      var fb = ((d << 10) | rem) ^ 0x5412;
      for (var a = 0; a <= 5; a++) setFn(8, a, bit(fb, a));
      setFn(8, 7, bit(fb, 6)); setFn(8, 8, bit(fb, 7)); setFn(7, 8, bit(fb, 8));
      for (var b = 9; b < 15; b++) setFn(14 - b, 8, bit(fb, b));
      for (var c = 0; c < 8; c++) setFn(size - 1 - c, 8, bit(fb, c));
      for (var e = 8; e < 15; e++) setFn(8, size - 15 + e, bit(fb, e));
      setFn(8, size - 8, true);
    }
    drawFormat(0);
    if (ver >= 7) {
      var r = ver;
      for (i = 0; i < 12; i++) r = (r << 1) ^ ((r >>> 11) * 0x1F25);
      var vb = (ver << 12) | r;
      for (i = 0; i < 18; i++) {
        var col = bit(vb, i), a2 = size - 11 + (i % 3), b2 = Math.floor(i / 3);
        setFn(a2, b2, col); setFn(b2, a2, col);
      }
    }

    // 데이터 비트 지그재그 배치
    var idx = 0;
    for (var right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right = 5;
      for (var vert = 0; vert < size; vert++) {
        for (j = 0; j < 2; j++) {
          var x = right - j, upward = ((right + 1) & 2) === 0, y = upward ? size - 1 - vert : vert;
          if (!isFn[y][x] && idx < all.length * 8) { modules[y][x] = bit(all[idx >>> 3], 7 - (idx & 7)); idx++; }
        }
      }
    }

    // 4) 마스크 8가지 중 벌점이 가장 낮은 것 고르기
    function applyMask(mask) {
      for (var y = 0; y < size; y++) for (var x = 0; x < size; x++) {
        var inv;
        switch (mask) {
          case 0: inv = (x + y) % 2 === 0; break;
          case 1: inv = y % 2 === 0; break;
          case 2: inv = x % 3 === 0; break;
          case 3: inv = (x + y) % 3 === 0; break;
          case 4: inv = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;
          case 5: inv = x * y % 2 + x * y % 3 === 0; break;
          case 6: inv = (x * y % 2 + x * y % 3) % 2 === 0; break;
          default: inv = ((x + y) % 2 + x * y % 3) % 2 === 0;
        }
        if (!isFn[y][x] && inv) modules[y][x] = !modules[y][x];
      }
    }
    function penalty() {
      var score = 0, y, x, run, c, dark = 0;
      for (y = 0; y < size; y++) {
        run = 1;
        for (x = 1; x < size; x++) {
          if (modules[y][x] === modules[y][x - 1]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;
        }
      }
      for (x = 0; x < size; x++) {
        run = 1;
        for (y = 1; y < size; y++) {
          if (modules[y][x] === modules[y - 1][x]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;
        }
      }
      for (y = 0; y < size - 1; y++) for (x = 0; x < size - 1; x++) {
        c = modules[y][x];
        if (c === modules[y][x + 1] && c === modules[y + 1][x] && c === modules[y + 1][x + 1]) score += 3;
      }
      var pat = [true, false, true, true, true, false, true, false, false, false, false];
      function matches(get) {
        var k2, ok1 = true, ok2 = true;
        for (k2 = 0; k2 < 11; k2++) { if (get(k2) !== pat[k2]) ok1 = false; if (get(k2) !== pat[10 - k2]) ok2 = false; }
        return ok1 || ok2;
      }
      for (y = 0; y < size; y++) for (x = 0; x <= size - 11; x++) { if (matches(function (k2) { return modules[y][x + k2]; })) score += 40; }
      for (x = 0; x < size; x++) for (y = 0; y <= size - 11; y++) { if (matches(function (k2) { return modules[y + k2][x]; })) score += 40; }
      for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (modules[y][x]) dark++;
      score += Math.floor(Math.abs(dark * 20 - size * size * 10) / (size * size)) * 10;
      return score;
    }
    var bestMask = 0, bestScore = Infinity, m;
    for (m = 0; m < 8; m++) {
      applyMask(m); drawFormat(m);
      var sc = penalty();
      if (sc < bestScore) { bestScore = sc; bestMask = m; }
      applyMask(m); // 되돌리기
    }
    applyMask(bestMask); drawFormat(bestMask);

    return { version: ver, size: size, modules: modules };
  }

  return { encode: encode };
})();
/* QR-END */
if (typeof module !== 'undefined') module.exports = QrCore;
