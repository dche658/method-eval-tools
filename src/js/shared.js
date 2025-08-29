/* Calculate the mean of a set of values
* arr must be an array of numbers.
*/
function average(arr) {
    if (arr.length === 0) {
        return 0; // Handle empty array case to avoid division by zero
    }
    let sum = 0;
    for (let i = 0; i < arr.length; i++) {
        sum += arr[i];
    }
    return sum / arr.length;
}

/*
* Sum of the sqaured deviations.
* 
* Sum of (x_i - mean(x))^2
*/
function devSq(arr) {
    const mean = average(arr);
    let sum = 0;
    for (let i = 0; i < arr.length; i++) {
        sum += (arr[i] - mean) * (arr[i] - mean);
    }
    return sum;
}

/* Calculate the standard deviation for a sample or population
*/
function stdev(arr, usePopulation = false) {
  const sumDevSq = devSq(arr);

  const variance = sumDevSq / 
                   (arr.length - (usePopulation ? 0 : 1));

  return Math.sqrt(variance);
};

/* Calculate the q th quantile for the given probabilities.
* 
* arr must be an array of numeric values
* probs is an array of probabilities ranging from 0 to 1
*/
function quantile(arr, probs) {
    if (!Array.isArray(arr) || arr.length === 0) {
        return undefined; // Handle empty or invalid input
    }

    for (let q of probs ) {
        if (q < 0 || q > 1) {
            throw new RangeError(`Quantiles must be between 0 and 1. q=${q}`);
        }
    }

    const sortedArr = [...arr].sort((a, b) => a - b); // Create a sorted copy
    const n = sortedArr.length;
    const quantiles = new Array(probs.length);

    let getQuantile = function(sortedArr, p) {
        if (p === 1) {
            return sortedArr[n - 1]; // Maximum value
        }
        if (p === 0) {
            return sortedArr[0]; // Minimum value
        }
        
        const index = (n - 1) * p;
        const lowerIndex = Math.floor(index);
        const upperIndex = Math.ceil(index);

        // Interpolate the value if the index is not an integer.
        if (lowerIndex === upperIndex) {
            return sortedArr[lowerIndex]; // Exact match
        } else {
            // Linear interpolation
            const lowerValue = sortedArr[lowerIndex];
            const upperValue = sortedArr[upperIndex];
            return lowerValue + (upperValue - lowerValue) * (index - lowerIndex);
        }
    }
    
    for (let i = 0; i < probs.length; i++) {
        quantiles[i] = getQuantile(sortedArr, probs[i]);
    }

    return quantiles;
}
