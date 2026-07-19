/**
 * KAMI Engine Bridge — optional WebGPU integration for 3D chart rendering.
 *
 * Unlike pptx.etzhayyim.com where KAMI renders slides, xlsx uses KAMI only for
 * 3D chart visualisation (bar, pie, surface, scatter). Spreadsheet grid
 * rendering uses Canvas 2D (grid-renderer.ts) as primary path.
 *
 * Falls back gracefully to Canvas 2D chart rendering if WebGPU/KAMI unavailable.
 */

let kamiModule: Record<string, unknown> | null = null;
let kamiLoading = false;

/** Check if WebGPU is available in the current browser. */
export async function checkWebGPU(): Promise<boolean> {
  const gpu = (navigator as Record<string, unknown>).gpu as { requestAdapter?: () => Promise<unknown> } | undefined;
  if (!gpu) return false;
  try {
    const adapter = await gpu.requestAdapter!();
    return adapter !== null;
  } catch {
    return false;
  }
}

/** Get GPU device info string. */
export async function getGPUInfo(): Promise<string | null> {
  const gpu = (navigator as Record<string, unknown>).gpu as { requestAdapter?: () => Promise<Record<string, unknown> | null> } | undefined;
  if (!gpu) return null;
  try {
    const adapter = await gpu.requestAdapter!();
    if (!adapter) return null;
    const info = await (adapter as Record<string, unknown>).requestAdapterInfo!() as Record<string, string>;
    return `${info.vendor} ${info.device} (${info.architecture})`;
  } catch {
    return null;
  }
}

/** Load the KAMI Engine WASM module (optional, non-blocking). */
export async function loadKamiEngine(): Promise<boolean> {
  if (kamiModule) return true;
  if (kamiLoading) return false;
  kamiLoading = true;

  try {
    const mod = await import(/* @vite-ignore */ "/pkg/kami_web.js");
    await mod.default();
    kamiModule = mod;
    return true;
  } catch {
    kamiLoading = false;
    return false;
  }
}

/** Check if KAMI Engine is loaded and ready. */
export function isKamiReady(): boolean {
  return kamiModule !== null;
}

/** Render a chart using KAMI Engine WebGPU (3D). */
export async function renderChartKami(
  canvasId: string,
  chartType: string,
  data: number[][],
  labels: string[],
  title: string,
): Promise<boolean> {
  if (!kamiModule) return false;

  try {
    const sceneJson = JSON.stringify({
      camera: {
        position: [0, 2, 5],
        target: [0, 0, 0],
        orthographic: false,
      },
      chart: {
        type: chartType,
        data,
        labels,
        title,
      },
    });

    const runWithScene = kamiModule.run_with_scene as (id: string, scene: string) => void;
    runWithScene(canvasId, sceneJson);
    return true;
  } catch {
    return false;
  }
}

/** Convert hex colour "#RRGGBB" to [r, g, b, a] normalised floats. */
export function hexToRgba(hex: string): [number, number, number, number] {
  const clean = hex.replace("#", "").replace(/^FF/i, "");
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return [r, g, b, 1.0];
}
