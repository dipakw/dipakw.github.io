import { colIndex, colName, numeric } from './cellref';
import type { CellValue } from './types';

/** A cell resolver: given (row, col) return its evaluated value. */
export type ResolveCell = (r: number, c: number) => CellValue;

type Token =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'id'; v: string }
  | { t: 'op'; v: string };

/** Lex a formula body (the part after '=') into tokens. */
export function tokenize(s: string): Token[] {
  const t: Token[] = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(s[i + 1] || ''))) {
      let j = i;
      while (j < s.length && /[0-9.]/.test(s[j])) j++;
      t.push({ t: 'num', v: parseFloat(s.slice(i, j)) });
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let str = '';
      while (j < s.length && s[j] !== '"') {
        str += s[j];
        j++;
      }
      t.push({ t: 'str', v: str });
      i = j + 1;
      continue;
    }
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i;
      while (j < s.length && /[A-Za-z0-9_$]/.test(s[j])) j++;
      t.push({ t: 'id', v: s.slice(i, j).replace(/\$/g, '') });
      i = j;
      continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      t.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^%(),:=<>&'.indexOf(ch) >= 0) {
      t.push({ t: 'op', v: ch });
      i++;
      continue;
    }
    throw { e: '#ERROR!' };
  }
  return t;
}

/**
 * Rewrite cell references in a formula body (the part after '=') when rows/cols
 * are inserted or deleted. `mapRow`/`mapCol` return the new index, or null if the
 * referenced row/col was removed (→ #REF!). String literals are left untouched.
 */
export function rewriteFormulaRefs(
  src: string,
  mapRow: (r: number) => number | null,
  mapCol: (c: number) => number | null,
): string {
  const refRe = /^\$?([A-Za-z]{1,2})\$?(\d{1,5})/;
  let out = '';
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      out += ch;
      i++;
      while (i < src.length && src[i] !== '"') {
        out += src[i];
        i++;
      }
      if (i < src.length) {
        out += src[i];
        i++;
      }
      continue;
    }
    const atBoundary = /[A-Za-z]/.test(ch) && (i === 0 || !/[A-Za-z0-9_$]/.test(src[i - 1]));
    if (atBoundary) {
      const m = refRe.exec(src.slice(i));
      if (m) {
        const after = src[i + m[0].length];
        // Not a ref if it runs into more identifier chars or a '(' (function call).
        if (!after || !/[A-Za-z0-9_$(]/.test(after)) {
          const nc = mapCol(colIndex(m[1]));
          const nr = mapRow(parseInt(m[2], 10) - 1);
          out += nc === null || nr === null ? '#REF!' : colName(nc) + (nr + 1);
          i += m[0].length;
          continue;
        }
      }
    }
    out += ch;
    i++;
  }
  return out;
}

/** Spreadsheet truthiness (used by IF / AND / OR / NOT). */
export function truthy(v: unknown): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === '' || v == null) return false;
  const n = numeric(v);
  if (n !== null) return n !== 0;
  return String(v).toUpperCase() !== 'FALSE';
}

/**
 * Evaluate a formula body (without the leading '=').
 *
 * `resolve` looks up other cells' values — the engine passes a closure that
 * threads its cycle-detection set, so ranges and refs recurse safely.
 *
 * Internally the evaluator juggles scalars and arrays (ranges), so the working
 * type is `any`; the public surface returns a single CellValue.
 */
export function evaluate(src: string, resolve: ResolveCell): CellValue {
  const tk = tokenize(src);
  let i = 0;
  const peek = () => tk[i];
  const eat = (v: string): boolean => {
    if (tk[i] && tk[i].t === 'op' && tk[i].v === v) {
      i++;
      return true;
    }
    return false;
  };

  const num = (v: any): number => {
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v === '' || v == null) return 0;
    if (typeof v === 'string' && v[0] === '#') throw { e: v };
    const n = numeric(v);
    if (n === null) throw { e: '#VALUE!' };
    return n;
  };

  const flat = (args: any[]): any[] => {
    const out: any[] = [];
    const walk = (a: any) => {
      if (Array.isArray(a)) a.forEach(walk);
      else out.push(a);
    };
    args.forEach(walk);
    return out;
  };
  const nums = (args: any[]): number[] => flat(args).filter((v) => v !== '' && v != null).map(num);

  const FN: Record<string, (a: any[]) => any> = {
    SUM: (a) => nums(a).reduce((x, y) => x + y, 0),
    PRODUCT: (a) => nums(a).reduce((x, y) => x * y, 1),
    AVERAGE: (a) => {
      const n = nums(a);
      if (!n.length) throw { e: '#DIV/0!' };
      return n.reduce((x, y) => x + y, 0) / n.length;
    },
    MIN: (a) => {
      const n = nums(a);
      return n.length ? Math.min.apply(null, n) : 0;
    },
    MAX: (a) => {
      const n = nums(a);
      return n.length ? Math.max.apply(null, n) : 0;
    },
    COUNT: (a) => nums(a).length,
    COUNTA: (a) => flat(a).filter((v) => v !== '' && v != null).length,
    MEDIAN: (a) => {
      const n = nums(a).sort((x, y) => x - y);
      if (!n.length) throw { e: '#DIV/0!' };
      const m = n.length >> 1;
      return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
    },
    ROUND: (a) => {
      const p = a.length > 1 ? num(a[1]) : 0;
      const f = Math.pow(10, p);
      return Math.round(num(a[0]) * f) / f;
    },
    ABS: (a) => Math.abs(num(a[0])),
    SQRT: (a) => Math.sqrt(num(a[0])),
    POWER: (a) => Math.pow(num(a[0]), num(a[1])),
    MOD: (a) => num(a[0]) % num(a[1]),
    INT: (a) => Math.floor(num(a[0])),
    IF: (a) => (truthy(a[0]) ? a[1] : a.length > 2 ? a[2] : false),
    AND: (a) => flat(a).every((v) => truthy(v)),
    OR: (a) => flat(a).some((v) => truthy(v)),
    NOT: (a) => !truthy(a[0]),
    LEN: (a) => String(a[0] == null ? '' : a[0]).length,
    UPPER: (a) => String(a[0] == null ? '' : a[0]).toUpperCase(),
    LOWER: (a) => String(a[0] == null ? '' : a[0]).toLowerCase(),
    CONCAT: (a) => flat(a).map((v) => (v == null ? '' : v)).join(''),
    TRIM: (a) => String(a[0] == null ? '' : a[0]).trim(),
    PI: () => Math.PI,
  };

  const cellRange = (a: { r: number; c: number }, b: { r: number; c: number }): any[] => {
    const out: any[] = [];
    for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++)
      for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) out.push(resolve(r, c));
    return out;
  };
  const refOf = (id: string): { r: number; c: number } | null => {
    const m = /^([A-Za-z]{1,2})(\d{1,5})$/.exec(id);
    return m ? { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) } : null;
  };

  const primary = (): any => {
    const t = peek();
    if (!t) throw { e: '#ERROR!' };
    if (t.t === 'num') {
      i++;
      return t.v;
    }
    if (t.t === 'str') {
      i++;
      return t.v;
    }
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) {
      i++;
      const v = primary();
      return t.v === '-' ? -num(v) : num(v);
    }
    if (t.t === 'op' && t.v === '(') {
      i++;
      const v = expr();
      if (!eat(')')) throw { e: '#ERROR!' };
      return v;
    }
    if (t.t === 'id') {
      i++;
      const up = t.v.toUpperCase();
      if (up === 'TRUE') return true;
      if (up === 'FALSE') return false;
      if (peek() && peek().t === 'op' && peek().v === '(') {
        i++;
        const args: any[] = [];
        if (!(peek() && peek().t === 'op' && peek().v === ')')) {
          args.push(expr());
          while (eat(',')) args.push(expr());
        }
        if (!eat(')')) throw { e: '#ERROR!' };
        const fn = FN[up] || (up === 'AVG' ? FN.AVERAGE : null);
        if (!fn) throw { e: '#NAME?' };
        return fn(args);
      }
      const a = refOf(t.v);
      if (!a) throw { e: '#NAME?' };
      if (peek() && peek().t === 'op' && peek().v === ':' && tk[i + 1] && tk[i + 1].t === 'id') {
        const b = refOf(tk[i + 1].v as string);
        if (b) {
          i += 2;
          return cellRange(a, b);
        }
      }
      return resolve(a.r, a.c);
    }
    throw { e: '#ERROR!' };
  };

  const pow = (): any => {
    const l = primary();
    if (eat('^')) return Math.pow(num(l), num(pow()));
    return l;
  };
  const mul = (): any => {
    let l = pow();
    while (peek() && peek().t === 'op' && (peek().v === '*' || peek().v === '/' || peek().v === '%')) {
      const op = tk[i++].v;
      const r = pow();
      if (op === '*') l = num(l) * num(r);
      else if (op === '/') {
        const d = num(r);
        if (d === 0) throw { e: '#DIV/0!' };
        l = num(l) / d;
      } else l = num(l) % num(r);
    }
    return l;
  };
  const add = (): any => {
    let l = mul();
    while (peek() && peek().t === 'op' && (peek().v === '+' || peek().v === '-' || peek().v === '&')) {
      const op = tk[i++].v;
      const r = mul();
      if (op === '+') l = num(l) + num(r);
      else if (op === '-') l = num(l) - num(r);
      else l = String(l == null ? '' : l) + String(r == null ? '' : r);
    }
    return l;
  };
  const expr = (): any => {
    let l = add();
    while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].indexOf(peek().v as string) >= 0) {
      const op = tk[i++].v;
      const r = add();
      const cmp =
        (typeof l === 'string' || typeof r === 'string') && numeric(l) === null && numeric(r) === null
          ? [String(l), String(r)]
          : [num(l), num(r)];
      if (op === '=') l = cmp[0] === cmp[1];
      else if (op === '<>') l = cmp[0] !== cmp[1];
      else if (op === '<') l = cmp[0] < cmp[1];
      else if (op === '>') l = cmp[0] > cmp[1];
      else if (op === '<=') l = cmp[0] <= cmp[1];
      else l = cmp[0] >= cmp[1];
    }
    return l;
  };

  const v = expr();
  if (i < tk.length) throw { e: '#ERROR!' };
  return Array.isArray(v) ? (v.length ? v[0] : '') : v;
}
