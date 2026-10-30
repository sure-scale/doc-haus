export const MODEL_CONCURRENCY_PRESETS = [1, 3, 5, 10, 20, 30, 50, 80, 100] as const
export const DEFAULT_MODEL_CONCURRENCY = 50

// Resolve ties upwards, so a previously saved limit of 4 becomes 5.
export function nearestModelConcurrency(value: number): number {
  return MODEL_CONCURRENCY_PRESETS.reduce((nearest, candidate) =>
    Math.abs(candidate - value) <= Math.abs(nearest - value) ? candidate : nearest,
  )
}
