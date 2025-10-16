/* This file is just used for debugging functions without having to 
 * bootstrap the whole add-in.
 */
import { stdev, studentt } from 'jstat-esm';


function grubbsTest(data, alpha = 0.01) {
  const outlier = {outlier: undefined, outlierIndex: undefined, g: undefined, gCrit: undefined}
  const n = data.length;
  const mean = data.reduce((a, b) => a + b, 0) / n;
  const t = studentt.inv(1 - alpha / (2 * n), n - 2);
  const gCrit = ((n - 1)/Math.sqrt(n))*Math.sqrt(Math.pow(t,2)/(n-2+Math.pow(t,2)));

  let maxDeltaIndex = 0;
  let maxDelta = 0;
  for (let i = 0; i < n; i++) {
    const delta = Math.abs(data[i] - mean);
    if (delta > maxDelta) {
      maxDelta = delta;
      maxDeltaIndex = i;
    }
  }
  const g = maxDelta / stdev(data);
  if (g > gCrit) {
    outlier.outlier = data[maxDeltaIndex];
    outlier.outlierIndex = maxDeltaIndex;
    outlier.g = g;
    outlier.gCrit = gCrit;
  }
  return outlier;
}

const data = [242,246,245,246,243,242,238,238,247,239,241,240,249,241,250,245,246,242,243,240,244,245,270,247,241];

const outlier = grubbsTest(data);
console.log(outlier);