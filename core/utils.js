export function tryWrap(fn, fallback = null) {
    try { return fn(); } catch (e) { return fallback; }
  }
  
  export function isFn(f) { return typeof f === 'function'; }
  