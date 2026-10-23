/**
 * Shape of the object returned by the density function,
 * mimicking the core properties of R's density object.
 */
interface DensityResult {
  /** The 512 equally spaced coordinates where the density is estimated. */
  x: number[];
  /** The estimated density values at the coordinates. */
  y: number[];
  /** The calculated bandwidth (smoothing parameter). */
  bw: number;
}

/**
 * Computes Kernel Density Estimation (KDE) equivalent to R's stats::density(x)
 * 
 * @param data - Array of numeric observations.
 * @param n - Number of equally spaced points at which the density is estimated (Default: 512).
 * @returns Object containing the grid points (x), density coordinates (y), and bandwidth (bw).
 */
export function density(data: number[], n: number = 512): DensityResult {
  if (!data || data.length === 0) {
    return { x: [], y: [], bw: 0 };
  }

  const len: number = data.length;

  // 1. Calculate descriptive statistics needed for Silverman's Rule (nrd0)
  const mean: number = data.reduce((sum, v) => sum + v, 0) / len;
  const variance: number = data.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / (len - 1);
  const sd: number = Math.sqrt(variance);

  // Calculate Interquartile Range (IQR) to handle skewed/bimodal distributions properly
  const sorted: number[] = [...data].sort((a, b) => a - b);
  const q25: number = sorted[Math.floor(len * 0.25)];
  const q75: number = sorted[Math.floor(len * 0.75)];
  const iqr: number = q75 - q25;

  // 2. Silverman's Rule of Thumb (R's default 'nrd0' bandwidth selection)
  // Formula: 0.9 * min(SD, IQR/1.34) * n^(-1/5)
  const minSigma: number = iqr > 0 ? Math.min(sd, iqr / 1.34) : sd;
  const bw: number = 0.9 * minSigma * Math.pow(len, -0.2);

  // 3. Define the Gaussian Kernel Function
  // K(x) = (1 / sqrt(2 * pi)) * exp(-0.5 * x^2)
  const SQRT_2PI: number = Math.sqrt(2 * Math.PI);
  const gaussianKernel = (x: number): number => {
    return Math.exp(-0.5 * x * x) / SQRT_2PI;
  };

  // 4. Generate the point grid range (mimicking R's default padding of 3 * bw)
  const minData: number = sorted[0];
  const maxData: number = sorted[len - 1];
  const from: number = minData - 3 * bw;
  const to: number = maxData + 3 * bw;
  const step: number = (to - from) / (n - 1);

  const xGrid: number[] = [];
  const yDensity: number[] = [];

  // 5. Evaluate the density at each grid point
  for (let i = 0; i < n; i++) {
    const x: number = from + i * step;
    
    // Average the kernel results for all observations relative to the current grid point
    let sum = 0;
    for (let j = 0; j < len; j++) {
      sum += gaussianKernel((x - data[j]) / bw);
    }
    
    xGrid.push(x);
    // Divide by (sample size * bandwidth) to ensure total area under curve equals 1
    yDensity.push(sum / (len * bw));
  }

  return { x: xGrid, y: yDensity, bw };
}
