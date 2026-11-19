import zlib from 'node:zlib';

function parseHexColor(hex: string, fallback: [number, number, number]): [number, number, number] {
  const clean = hex.replace('#', '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(clean)) return fallback;
  return [
    parseInt(clean.slice(0, 2), 16),
    parseInt(clean.slice(2, 4), 16),
    parseInt(clean.slice(4, 6), 16),
  ];
}

function crc32(buf: Buffer): number {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function makeChunk(type: string, data: Buffer): Buffer {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcInput = Buffer.concat([typeBuf, data]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(crcInput), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function mixColor(
  a: [number, number, number],
  b: [number, number, number],
  t: number
): [number, number, number] {
  return [
    Math.round(a[0] + (b[0] - a[0]) * t),
    Math.round(a[1] + (b[1] - a[1]) * t),
    Math.round(a[2] + (b[2] - a[2]) * t),
  ];
}

function inRoundedRect(
  x: number,
  y: number,
  rx: number,
  ry: number,
  rw: number,
  rh: number,
  radius: number
): boolean {
  if (x < rx || x > rx + rw || y < ry || y > ry + rh) return false;
  const cx = Math.max(rx + radius, Math.min(x, rx + rw - radius));
  const cy = Math.max(ry + radius, Math.min(y, ry + rh - radius));
  const dx = x - cx;
  const dy = y - cy;
  return dx * dx + dy * dy <= radius * radius;
}

function inRotatedEllipse(
  x: number,
  y: number,
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  deg: number
): boolean {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = x - cx;
  const dy = y - cy;
  const u = dx * cos + dy * sin;
  const v = -dx * sin + dy * cos;
  return (u * u) / (rx * rx) + (v * v) / (ry * ry) <= 1;
}

function distToSegment(
  px: number,
  py: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number
): number {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const lenSq = dx * dx + dy * dy;
  if (lenSq === 0) return Math.hypot(px - x1, py - y1);
  const t = Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / lenSq));
  return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

function distToQuadBezier(
  px: number,
  py: number,
  x0: number,
  y0: number,
  cx: number,
  cy: number,
  x1: number,
  y1: number
): number {
  let minD = Infinity;
  let prevX = x0;
  let prevY = y0;
  const steps = 16;
  for (let i = 1; i <= steps; i++) {
    const t = i / steps;
    const mt = 1 - t;
    const bx = mt * mt * x0 + 2 * mt * t * cx + t * t * x1;
    const by = mt * mt * y0 + 2 * mt * t * cy + t * t * y1;
    const d = distToSegment(px, py, prevX, prevY, bx, by);
    if (d < minD) minD = d;
    prevX = bx;
    prevY = by;
  }
  return minD;
}

interface GrainSpec {
  cx: number;
  cy: number;
  rx: number;
  ry: number;
  deg: number;
  color: [number, number, number];
}

const GOLD_TOP: [number, number, number] = [254, 240, 138]; // #FEF08A
const GOLD_1: [number, number, number] = [253, 224, 71]; // #FDE047
const GOLD_2: [number, number, number] = [250, 204, 21]; // #FACC15
const GOLD_3: [number, number, number] = [251, 191, 36]; // #FBBF24
const GOLD_4: [number, number, number] = [245, 158, 11]; // #F59E0B
const STEM_SIDE: [number, number, number] = [252, 211, 77]; // #FCD34D

const GRAINS: GrainSpec[] = [
  // Left stalk (in 0..1 normalized from 512x512)
  { cx: 154 / 512, cy: 156 / 512, rx: 13 / 512, ry: 23 / 512, deg: -22, color: GOLD_1 },
  { cx: 140 / 512, cy: 198 / 512, rx: 13 / 512, ry: 22 / 512, deg: -50, color: GOLD_2 },
  { cx: 178 / 512, cy: 186 / 512, rx: 13 / 512, ry: 22 / 512, deg: 12, color: GOLD_2 },
  { cx: 150 / 512, cy: 244 / 512, rx: 13.5 / 512, ry: 22 / 512, deg: -50, color: GOLD_3 },
  { cx: 190 / 512, cy: 232 / 512, rx: 13.5 / 512, ry: 22 / 512, deg: 12, color: GOLD_3 },
  { cx: 166 / 512, cy: 290 / 512, rx: 13 / 512, ry: 21 / 512, deg: -48, color: GOLD_4 },
  { cx: 206 / 512, cy: 278 / 512, rx: 13 / 512, ry: 21 / 512, deg: 14, color: GOLD_4 },

  // Right stalk
  { cx: 358 / 512, cy: 156 / 512, rx: 13 / 512, ry: 23 / 512, deg: 22, color: GOLD_1 },
  { cx: 372 / 512, cy: 198 / 512, rx: 13 / 512, ry: 22 / 512, deg: 50, color: GOLD_2 },
  { cx: 334 / 512, cy: 186 / 512, rx: 13 / 512, ry: 22 / 512, deg: -12, color: GOLD_2 },
  { cx: 362 / 512, cy: 244 / 512, rx: 13.5 / 512, ry: 22 / 512, deg: 50, color: GOLD_3 },
  { cx: 322 / 512, cy: 232 / 512, rx: 13.5 / 512, ry: 22 / 512, deg: -12, color: GOLD_3 },
  { cx: 346 / 512, cy: 290 / 512, rx: 13 / 512, ry: 21 / 512, deg: 48, color: GOLD_4 },
  { cx: 306 / 512, cy: 278 / 512, rx: 13 / 512, ry: 21 / 512, deg: -14, color: GOLD_4 },

  // Center main stalk
  { cx: 256 / 512, cy: 112 / 512, rx: 16 / 512, ry: 28 / 512, deg: 0, color: GOLD_TOP },
  { cx: 230 / 512, cy: 156 / 512, rx: 15 / 512, ry: 27 / 512, deg: -36, color: GOLD_1 },
  { cx: 282 / 512, cy: 156 / 512, rx: 15 / 512, ry: 27 / 512, deg: 36, color: GOLD_1 },
  { cx: 226 / 512, cy: 204 / 512, rx: 16 / 512, ry: 28 / 512, deg: -38, color: GOLD_2 },
  { cx: 286 / 512, cy: 204 / 512, rx: 16 / 512, ry: 28 / 512, deg: 38, color: GOLD_2 },
  { cx: 226 / 512, cy: 254 / 512, rx: 16 / 512, ry: 28 / 512, deg: -38, color: GOLD_3 },
  { cx: 286 / 512, cy: 254 / 512, rx: 16 / 512, ry: 28 / 512, deg: 38, color: GOLD_3 },
  { cx: 229 / 512, cy: 304 / 512, rx: 15 / 512, ry: 26 / 512, deg: -36, color: GOLD_4 },
  { cx: 283 / 512, cy: 304 / 512, rx: 15 / 512, ry: 26 / 512, deg: 36, color: GOLD_4 },
];

function inSweepingLeaf(gx: number, gy: number): boolean {
  // Two graceful crescent leaves cradling the bottom of the sheaf
  const inLeftOuter = inRotatedEllipse(gx, gy, 0.355, 0.69, 0.165, 0.068, 34);
  const inLeftCut = inRotatedEllipse(gx, gy, 0.39, 0.655, 0.165, 0.068, 34);
  const inLeft = inLeftOuter && !inLeftCut && gx >= 0.22 && gx <= 0.5 && gy >= 0.54 && gy <= 0.82;

  const inRightOuter = inRotatedEllipse(gx, gy, 0.645, 0.69, 0.165, 0.068, -34);
  const inRightCut = inRotatedEllipse(gx, gy, 0.61, 0.655, 0.165, 0.068, -34);
  const inRight = inRightOuter && !inRightCut && gx >= 0.5 && gx <= 0.78 && gy >= 0.54 && gy <= 0.82;

  return inLeft || inRight;
}

export interface PngIconOptions {
  size: number;
  primaryColor?: string;
  secondaryColor?: string;
  maskable?: boolean;
  roundedBackground?: boolean;
}

/**
 * Pure Node.js RGBA PNG generator for PWA & Android mipmap icons.
 * Renders a crisp golden paddy/wheat sheaf food-grain emblem with 2x2 supersampling.
 */
export function generateFoodGrainPngIcon(options: PngIconOptions): Buffer {
  const {
    size,
    primaryColor = '#0D6E4E',
    secondaryColor = '#059669',
    maskable = false,
    roundedBackground = true,
  } = options;

  const c1 = parseHexColor(primaryColor, [13, 110, 78]);
  const c2 = parseHexColor(secondaryColor, [5, 150, 105]);
  const ringBg = mixColor(c1, [6, 95, 70], 0.55);
  const ringStroke: [number, number, number] = [110, 231, 183];
  const white: [number, number, number] = [255, 255, 255];

  const padScale = maskable ? 0.74 : 0.92;
  const raw = Buffer.alloc(size * (1 + size * 4));
  const subOffsets = [0.25, 0.75];

  for (let y = 0; y < size; y++) {
    const rowStart = y * (1 + size * 4);
    raw[rowStart] = 0;

    for (let x = 0; x < size; x++) {
      let accR = 0;
      let accG = 0;
      let accB = 0;
      let accA = 0;

      for (const oy of subOffsets) {
        for (const ox of subOffsets) {
          const nx = (x + ox) / size;
          const ny = (y + oy) / size;

          const bgActive =
            maskable || !roundedBackground
              ? true
              : inRoundedRect(nx, ny, 0, 0, 1, 1, 0.22);

          if (!bgActive) continue;

          const gradT = Math.min(1, Math.max(0, (nx + ny) * 0.5));
          let [r, g, b] = mixColor(c1, c2, gradT);

          const gx = (nx - 0.5) / padScale + 0.5;
          const gy = (ny - 0.5) / padScale + 0.5;

          const distCenter = Math.hypot(gx - 0.5, gy - 0.5);
          if (distCenter <= 196 / 512) {
            if (distCenter >= 192 / 512) {
              [r, g, b] = mixColor([r, g, b], ringStroke, 0.3);
            } else {
              [r, g, b] = mixColor([r, g, b], ringBg, 0.34);
            }
          }

          // Side curved stems
          const dLeftStem = distToQuadBezier(
            gx,
            gy,
            250 / 512,
            398 / 512,
            186 / 512,
            290 / 512,
            162 / 512,
            178 / 512
          );
          const dRightStem = distToQuadBezier(
            gx,
            gy,
            262 / 512,
            398 / 512,
            326 / 512,
            290 / 512,
            350 / 512,
            178 / 512
          );
          if (dLeftStem <= 4 / 512 || dRightStem <= 4 / 512) {
            [r, g, b] = STEM_SIDE;
          }

          // Center upright stem
          const dCenterStem = distToSegment(gx, gy, 0.5, 416 / 512, 0.5, 146 / 512);
          if (dCenterStem <= 5 / 512) {
            [r, g, b] = GOLD_TOP;
          }

          // Grain kernels
          for (const gr of GRAINS) {
            if (inRotatedEllipse(gx, gy, gr.cx, gr.cy, gr.rx, gr.ry, gr.deg)) {
              [r, g, b] = gr.color;
              break;
            }
          }

          // Sweeping white leaves & binding band
          if (
            inSweepingLeaf(gx, gy) ||
            inRoundedRect(gx, gy, 234 / 512, 384 / 512, 44 / 512, 16 / 512, 8 / 512)
          ) {
            [r, g, b] = white;
          }

          accR += r;
          accG += g;
          accB += b;
          accA += 255;
        }
      }

      const px = rowStart + 1 + x * 4;
      if (accA === 0) {
        raw[px] = 0;
        raw[px + 1] = 0;
        raw[px + 2] = 0;
        raw[px + 3] = 0;
      } else {
        const samples = accA / 255;
        raw[px] = Math.round(accR / samples);
        raw[px + 1] = Math.round(accG / samples);
        raw[px + 2] = Math.round(accB / samples);
        raw[px + 3] = Math.round(accA / 4);
      }
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const idat = zlib.deflateSync(raw, { level: 6 });
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  return Buffer.concat([
    signature,
    makeChunk('IHDR', ihdr),
    makeChunk('IDAT', idat),
    makeChunk('IEND', Buffer.alloc(0)),
  ]);
}

export const generateChallanPngIcon = generateFoodGrainPngIcon;
