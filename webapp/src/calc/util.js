// Small helpers shared by the calculation modules. A value that cannot be worked out is `undefined`
// (the workbook shows "" there); a reason text goes into a separate field, never into a number.

export const isNum = x => typeof x === "number" && Number.isFinite(x);
// Excel N(): a number stays, anything else is 0.
export const N = x => (isNum(x) ? x : 0);
export const sum = xs => xs.reduce((a, x) => a + N(x), 0);
export const pos = x => isNum(x) && x > 0;

// Excel TEXT(x, "0.0") style: rounds half away from zero.
export function fixed(x, d) {
  const f = 10 ** d;
  const r = Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f;
  return (Object.is(r, -0) ? 0 : r).toFixed(d);
}
export const percent = x => fixed(x * 100, 0) + "%";

// Number as Excel writes it when joined to text ("15", "0.8").
export const numText = x => String(+x.toPrecision(15));

// Rank from the top, ties broken by order, as COUNTIF(>x) + COUNTIF(above and this, x). Items without a
// number come after all numbered ones, in order.
export function ranks(values) {
  const nums = values.filter(isNum);
  let blank = 0;
  return values.map((v, i) => {
    if (!isNum(v)) { blank++; return nums.length + blank; }
    return nums.filter(x => x > v).length + values.slice(0, i + 1).filter(x => x === v).length;
  });
}

// Put k: v on o only when v is defined (blank values are left out of the files).
export function put(o, k, v) { if (v !== undefined && v !== "") o[k] = v; return o; }
