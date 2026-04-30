declare const CustomFunctions: {
  Error: new (code: string, message?: string) => Error;
  ErrorCode: Record<string, string>;
  associate(id: string, fn: Function): void;
};

import { jStat } from "jstat";

type Primitive = string | number | boolean | null | undefined;
type ExcelInput = Primitive | Primitive[] | Primitive[][];
type ExcelOutput = Primitive | Primitive[][];
type HeaderMode = 0 | 1 | 2;
type Direction = "two" | "left" | "right";

const ADDIN_BRAND = "Evalytics";
const ADDIN_RUNTIME = "Office.js";
const ADDIN_VERSION = "1.0.0";
declare const __EVALYTICS_BUILD_STAMP__: string;

const ADDIN_BUILD = __EVALYTICS_BUILD_STAMP__;

const HEADER_AUTO: HeaderMode = 0;
const HEADER_HAS: HeaderMode = 1;

function invalidValue(message = "Invalid value"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
}

function invalidNumber(message = "Invalid number"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber, message);
}

function notAvailable(message = "Not available"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, message);
}

function divisionByZero(message = "Division by zero"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.divisionByZero, message);
}

function isBlank(value: Primitive): boolean {
  return value === null || value === undefined || (typeof value === "string" && value.trim().length === 0);
}

function flatten(input: ExcelInput): Primitive[] {
  if (Array.isArray(input)) {
    if (input.length === 0) {
      return [];
    }

    if (Array.isArray(input[0])) {
      return (input as Primitive[][]).flat();
    }

    return input as Primitive[];
  }

  return [input];
}

function tryGetNumber(value: Primitive): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  if (typeof value === "string") {
    const normalized = value.trim().replace(",", ".");
    if (!normalized) {
      return null;
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function shouldSkipLeadingHeader(items: Primitive[], dataPredicate: (value: Primitive) => boolean): boolean {
  if (items.length < 2) {
    return false;
  }

  const first = items[0];
  if (isBlank(first) || dataPredicate(first)) {
    return false;
  }

  let hasDataAfterFirst = false;
  for (let i = 1; i < items.length; i += 1) {
    const item = items[i];
    if (isBlank(item)) {
      continue;
    }

    hasDataAfterFirst = true;
    if (!dataPredicate(item)) {
      return false;
    }
  }

  return hasDataAfterFirst;
}

function skipLeadingHeader(items: Primitive[], dataPredicate: (value: Primitive) => boolean): Primitive[] {
  return shouldSkipLeadingHeader(items, dataPredicate) ? items.slice(1) : items;
}

function parseHeaderMode(input?: Primitive): HeaderMode {
  if (input === null || input === undefined || input === "") {
    return HEADER_AUTO;
  }

  const value = tryGetNumber(input);
  if (value === 0 || value === 1 || value === 2) {
    return value as HeaderMode;
  }

  throw invalidValue("Invalid hasHeader flag");
}

function parseDirection(input?: Primitive): Direction {
  if (input === null || input === undefined || input === "") {
    return "two";
  }

  const code = tryGetNumber(input);
  if (code === 0) {
    return "two";
  }
  if (code === 1) {
    return "left";
  }
  if (code === 2) {
    return "right";
  }

  throw invalidValue("Invalid direction");
}

function validateAlpha(alpha: number): void {
  if (!(alpha > 0 && alpha < 1)) {
    throw invalidNumber("Alpha must be between 0 and 1");
  }
}

function readNumericVector(input: ExcelInput, headerMode: HeaderMode = HEADER_AUTO): number[] {
  let items = flatten(input);
  if (headerMode === HEADER_HAS && items.length > 0) {
    items = items.slice(1);
  } else if (headerMode === HEADER_AUTO) {
    items = skipLeadingHeader(items, (value) => tryGetNumber(value) !== null);
  }

  const values: number[] = [];
  for (const item of items) {
    if (isBlank(item)) {
      continue;
    }

    const parsed = tryGetNumber(item);
    if (parsed === null) {
      throw invalidValue("Expected numeric data");
    }

    values.push(parsed);
  }

  return values;
}

function readPairedNumericWeights(valuesInput: ExcelInput, weightsInput: ExcelInput, headerMode: HeaderMode = HEADER_AUTO): [number[], number[]] {
  let rawValues = flatten(valuesInput);
  let rawWeights = flatten(weightsInput);

  if (rawValues.length !== rawWeights.length) {
    throw invalidValue("Values and weights must have the same length");
  }

  if (headerMode === HEADER_HAS && rawValues.length > 0) {
    rawValues = rawValues.slice(1);
    rawWeights = rawWeights.slice(1);
  } else if (
    headerMode === HEADER_AUTO &&
    shouldSkipLeadingHeader(rawValues, (item) => tryGetNumber(item) !== null) &&
    shouldSkipLeadingHeader(rawWeights, (item) => isBlank(item) || tryGetNumber(item) !== null)
  ) {
    rawValues = rawValues.slice(1);
    rawWeights = rawWeights.slice(1);
  }

  const values: number[] = [];
  const weights: number[] = [];

  for (let i = 0; i < rawValues.length; i += 1) {
    if (isBlank(rawValues[i])) {
      continue;
    }

    const value = tryGetNumber(rawValues[i]);
    if (value === null) {
      throw invalidValue("Expected numeric value");
    }

    let weight = 0;
    if (!isBlank(rawWeights[i])) {
      const parsedWeight = tryGetNumber(rawWeights[i]);
      if (parsedWeight === null) {
        throw invalidValue("Expected numeric weight");
      }
      weight = parsedWeight;
    }

    values.push(value);
    weights.push(weight);
  }

  return [values, weights];
}

function readPairedNumericVectors(firstInput: ExcelInput, secondInput: ExcelInput, headerMode: HeaderMode = HEADER_AUTO): [number[], number[]] {
  let rawFirst = flatten(firstInput);
  let rawSecond = flatten(secondInput);

  if (rawFirst.length !== rawSecond.length) {
    throw invalidValue("Paired ranges must have the same length");
  }

  if (headerMode === HEADER_HAS && rawFirst.length > 0) {
    rawFirst = rawFirst.slice(1);
    rawSecond = rawSecond.slice(1);
  } else if (
    headerMode === HEADER_AUTO &&
    shouldSkipLeadingHeader(rawFirst, (item) => tryGetNumber(item) !== null) &&
    shouldSkipLeadingHeader(rawSecond, (item) => tryGetNumber(item) !== null)
  ) {
    rawFirst = rawFirst.slice(1);
    rawSecond = rawSecond.slice(1);
  }

  const first: number[] = [];
  const second: number[] = [];

  for (let i = 0; i < rawFirst.length; i += 1) {
    if (isBlank(rawFirst[i]) || isBlank(rawSecond[i])) {
      continue;
    }

    const left = tryGetNumber(rawFirst[i]);
    const right = tryGetNumber(rawSecond[i]);
    if (left === null || right === null) {
      throw invalidValue("Expected paired numeric data");
    }

    first.push(left);
    second.push(right);
  }

  return [first, second];
}

function readBinaryVector(input: ExcelInput, headerMode: HeaderMode = HEADER_AUTO): [number, number] {
  let items = flatten(input);
  if (headerMode === HEADER_HAS && items.length > 0) {
    items = items.slice(1);
  } else if (headerMode === HEADER_AUTO) {
    items = skipLeadingHeader(items, (value) => value === true || value === false || tryGetNumber(value) === 0 || tryGetNumber(value) === 1);
  }

  let successes = 0;
  let count = 0;

  for (const item of items) {
    if (isBlank(item)) {
      continue;
    }

    if (typeof item === "boolean") {
      count += 1;
      successes += item ? 1 : 0;
      continue;
    }

    const parsed = tryGetNumber(item);
    if (parsed === null || (parsed !== 0 && parsed !== 1)) {
      throw invalidValue("Expected binary data");
    }

    count += 1;
    successes += parsed;
  }

  return [successes, count];
}

function readGroups(categoriesInput: ExcelInput, valuesInput: ExcelInput, headerMode: HeaderMode = HEADER_AUTO): Map<string, number[]> {
  let rawCategories = flatten(categoriesInput);
  let rawValues = flatten(valuesInput);

  if (rawCategories.length !== rawValues.length) {
    throw invalidValue("Categories and values must have the same length");
  }

  if (headerMode === HEADER_HAS && rawCategories.length > 0) {
    rawCategories = rawCategories.slice(1);
    rawValues = rawValues.slice(1);
  } else if (
    headerMode === HEADER_AUTO &&
    !isBlank(rawCategories[0]) &&
    tryGetNumber(rawValues[0]) === null &&
    shouldSkipLeadingHeader(rawValues, (item) => tryGetNumber(item) !== null)
  ) {
    rawCategories = rawCategories.slice(1);
    rawValues = rawValues.slice(1);
  }

  const groups = new Map<string, number[]>();
  for (let i = 0; i < rawCategories.length; i += 1) {
    if (isBlank(rawCategories[i]) || isBlank(rawValues[i])) {
      continue;
    }

    const value = tryGetNumber(rawValues[i]);
    if (value === null) {
      throw invalidValue("Expected numeric group value");
    }

    const key = String(rawCategories[i]);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(value);
  }

  return groups;
}

function mean(values: number[]): number {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function sampleVariance(values: number[]): number {
  const avg = mean(values);
  return values.reduce((sum, value) => sum + (value - avg) ** 2, 0) / (values.length - 1);
}

function populationVariance(values: number[]): number {
  const avg = mean(values);
  return values.reduce((sum, value) => sum + (value - avg) ** 2, 0) / values.length;
}

function sampleStandardDeviation(values: number[]): number {
  return Math.sqrt(sampleVariance(values));
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function pearsonCorrelation(x: number[], y: number[]): number {
  const meanX = mean(x);
  const meanY = mean(y);
  let sumXY = 0;
  let sumXX = 0;
  let sumYY = 0;

  for (let i = 0; i < x.length; i += 1) {
    const dx = x[i] - meanX;
    const dy = y[i] - meanY;
    sumXY += dx * dy;
    sumXX += dx * dx;
    sumYY += dy * dy;
  }

  return sumXY / Math.sqrt(sumXX * sumYY);
}

function midRank(values: number[]): number[] {
  const indexed = values.map((value, index) => ({ value, index })).sort((a, b) => a.value - b.value);
  const ranks = new Array(values.length).fill(0);
  let i = 0;

  while (i < indexed.length) {
    let j = i + 1;
    while (j < indexed.length && indexed[j].value === indexed[i].value) {
      j += 1;
    }

    const averageRank = ((i + 1) + j) / 2;
    for (let k = i; k < j; k += 1) {
      ranks[indexed[k].index] = averageRank;
    }

    i = j;
  }

  return ranks;
}

function validateWeights(values: number[], weights: number[]): void {
  if (values.length !== weights.length) {
    throw invalidValue("Values and weights must have the same length");
  }
  if (weights.some((weight) => weight < 0)) {
    throw invalidNumber("Weights must be non-negative");
  }
  if (weights.reduce((sum, value) => sum + value, 0) <= 0) {
    throw invalidNumber("Weights must sum to a positive value");
  }
}

function weightedMean(values: number[], weights: number[]): number {
  const sumWeights = weights.reduce((sum, value) => sum + value, 0);
  return values.reduce((sum, value, index) => sum + value * weights[index], 0) / sumWeights;
}

function weightedVariance(values: number[], weights: number[], sample: boolean): number {
  validateWeights(values, weights);
  const sumWeights = weights.reduce((sum, value) => sum + value, 0);
  if (sample && sumWeights <= 1) {
    throw invalidValue("Not enough weighted observations");
  }
  const avg = weightedMean(values, weights);
  const sumSquares = values.reduce((sum, value, index) => sum + weights[index] * (value - avg) ** 2, 0);
  return sumSquares / (sample ? sumWeights - 1 : sumWeights);
}

function criticalT(alpha: number, df: number, direction: Direction): number {
  return Math.abs(jStat.studentt.inv(direction === "two" ? 1 - (alpha / 2) : 1 - alpha, df));
}

function criticalZ(alpha: number, direction: Direction): number {
  return Math.abs(jStat.normal.inv(direction === "two" ? 1 - (alpha / 2) : 1 - alpha, 0, 1));
}

function pValueFromT(statistic: number, df: number, direction: Direction): number {
  const cdf = jStat.studentt.cdf(statistic, df);
  if (direction === "left") {
    return cdf;
  }
  if (direction === "right") {
    return 1 - cdf;
  }
  return 2 * (1 - jStat.studentt.cdf(Math.abs(statistic), df));
}

function pValueFromZ(statistic: number, direction: Direction): number {
  const cdf = jStat.normal.cdf(statistic, 0, 1);
  if (direction === "left") {
    return cdf;
  }
  if (direction === "right") {
    return 1 - cdf;
  }
  return 2 * (1 - jStat.normal.cdf(Math.abs(statistic), 0, 1));
}

function criticalLabel(symbol: string, direction: Direction): string {
  return direction === "two" ? `${symbol}ᶜʳⁱᵗ(1−α/2)` : `${symbol}ᶜʳⁱᵗ(1−α)`;
}

function buildRows(rows: Primitive[][]): Primitive[][] {
  return rows;
}

function rectangularRows(rows: Primitive[][], width?: number): Primitive[][] {
  const columnCount = width ?? rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
  return rows.map((row) => [...row, ...new Array(Math.max(0, columnCount - row.length)).fill("")]);
}

function percentileInc(sortedValues: number[], quantile: number): number {
  if (sortedValues.length === 1) {
    return sortedValues[0];
  }
  const position = quantile * (sortedValues.length - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) {
    return sortedValues[lower];
  }
  const fraction = position - lower;
  return sortedValues[lower] + fraction * (sortedValues[upper] - sortedValues[lower]);
}

function percentileExc(sortedValues: number[], quantile: number): number {
  const position = (quantile * (sortedValues.length + 1)) - 1;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) {
    return sortedValues[lower];
  }
  const fraction = position - lower;
  return sortedValues[lower] + fraction * (sortedValues[upper] - sortedValues[lower]);
}

function matchesCriterion(actual: Primitive, criterion: Primitive): boolean {
  if (criterion === null || criterion === undefined) {
    return actual === criterion;
  }

  if (typeof criterion === "string") {
    const operators = [">=", "<=", "<>", ">", "<", "="];
    let op = "=";
    let operand = criterion;
    for (const candidate of operators) {
      if (criterion.startsWith(candidate)) {
        op = candidate;
        operand = criterion.slice(candidate.length);
        break;
      }
    }
    return compareValues(actual, op, operand);
  }

  return compareValues(actual, "=", criterion);
}

function compareValues(actual: Primitive, operator: string, operand: Primitive): boolean {
  const actualNumber = tryGetNumber(actual);
  const operandNumber = tryGetNumber(operand);
  if (actualNumber !== null && operandNumber !== null) {
    switch (operator) {
      case "=": return actualNumber === operandNumber;
      case "<>": return actualNumber !== operandNumber;
      case ">": return actualNumber > operandNumber;
      case "<": return actualNumber < operandNumber;
      case ">=": return actualNumber >= operandNumber;
      case "<=": return actualNumber <= operandNumber;
      default: return false;
    }
  }

  const actualText = isBlank(actual) ? "" : String(actual);
  const operandText = String(operand ?? "");
  if (operandText.includes("*") || operandText.includes("?")) {
    const regex = new RegExp(`^${operandText.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "i");
    const matched = regex.test(actualText);
    return operator === "<>" ? !matched : matched;
  }

  const comparison = actualText.localeCompare(operandText, undefined, { sensitivity: "accent" });
  switch (operator) {
    case "=": return comparison === 0;
    case "<>": return comparison !== 0;
    case ">": return comparison > 0;
    case "<": return comparison < 0;
    case ">=": return comparison >= 0;
    case "<=": return comparison <= 0;
    default: return false;
  }
}

function applyFilters(valuesInput: ExcelInput, criteriaArgs: Primitive[]): number[] {
  const rawValues = flatten(valuesInput);
  if (criteriaArgs.length % 2 !== 0) {
    throw invalidValue("Criteria arguments must be in range/criterion pairs");
  }

  const include = new Array(rawValues.length).fill(true);
  for (let i = 0; i < criteriaArgs.length; i += 2) {
    const range = flatten(criteriaArgs[i] as ExcelInput);
    if (range.length !== rawValues.length) {
      throw invalidValue("Criteria range length mismatch");
    }
    const criterion = criteriaArgs[i + 1];
    for (let j = 0; j < range.length; j += 1) {
      include[j] = include[j] && matchesCriterion(range[j], criterion);
    }
  }

  const filtered: number[] = [];
  for (let i = 0; i < rawValues.length; i += 1) {
    if (!include[i] || isBlank(rawValues[i])) {
      continue;
    }
    const parsed = tryGetNumber(rawValues[i]);
    if (parsed === null) {
      throw invalidValue("Expected numeric value");
    }
    filtered.push(parsed);
  }

  if (filtered.length === 0) {
    throw notAvailable("No values matched the filters");
  }

  return filtered;
}

function kolmogorovComplementaryCdf(lambda: number): number {
  if (lambda <= 0) {
    return 1;
  }
  let sum = 0;
  for (let j = 1; j < 100; j += 1) {
    const term = Math.exp(-2 * j * j * lambda * lambda);
    sum += (j % 2 === 1 ? 1 : -1) * term;
    if (term < 1e-12) {
      break;
    }
  }
  return Math.min(1, Math.max(0, 2 * sum));
}

function weibullShapeEstimate(values: number[]): number {
  const logMean = mean(values.map((value) => Math.log(value)));
  let shape = 1.2 / sampleStandardDeviation(values.map((value) => Math.log(value)));
  shape = Number.isFinite(shape) && shape > 0 ? shape : 1;

  for (let iteration = 0; iteration < 100; iteration += 1) {
    let sumXk = 0;
    let sumXkLogX = 0;
    let sumXkLogXSq = 0;
    for (const value of values) {
      const xk = value ** shape;
      const logX = Math.log(value);
      sumXk += xk;
      sumXkLogX += xk * logX;
      sumXkLogXSq += xk * logX * logX;
    }
    const weightedLogMean = sumXkLogX / sumXk;
    const fn = (1 / shape) + logMean - weightedLogMean;
    const derivative = (-1 / (shape * shape)) - ((sumXkLogXSq / sumXk) - (weightedLogMean * weightedLogMean));
    let next = shape - (fn / derivative);
    if (!Number.isFinite(next) || next <= 0) {
      next = shape / 2;
    }
    if (Math.abs(next - shape) < 1e-10) {
      return next;
    }
    shape = next;
  }

  return shape;
}

function weibullScaleEstimate(values: number[], shape: number): number {
  return mean(values.map((value) => value ** shape)) ** (1 / shape);
}

function normalSample(meanValue: number, standardDeviation: number): number {
  const u1 = Math.max(Number.MIN_VALUE, Math.random());
  const u2 = Math.random();
  const standard = Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
  return meanValue + (standardDeviation * standard);
}

function parseOptionalProbability(input: Primitive, defaultValue = 0): number {
  if (isBlank(input)) {
    return defaultValue;
  }
  const result = tryGetNumber(input);
  if (result === null || result < 0 || result > 1) {
    throw invalidValue("Probability must be between 0 and 1");
  }
  return result;
}

function parseOptionalInteger(input: Primitive, defaultValue: number): number {
  if (isBlank(input)) {
    return defaultValue;
  }
  const parsed = tryGetNumber(input);
  if (parsed === null || Math.abs(parsed - Math.round(parsed)) > 1e-9) {
    throw invalidValue("Expected integer");
  }
  return Math.round(parsed);
}

function fill(what: string, count: Primitive): string[][] {
  const repeat = parseOptionalInteger(count, 0);
  if (repeat < 1) {
    throw invalidNumber("Count must be >= 1");
  }
  const value = what ?? "";
  const values: string[] = [];
  for (let index = 0; index < repeat; index += 1) {
    values.push(value);
  }

  return values.map((value) => [value]);
}

function runtimeDecimalSeparator(): "." | "," {
  try {
    const decimalPart = new Intl.NumberFormat(undefined).formatToParts(1.1).find((part) => part.type === "decimal");
    return decimalPart?.value === "," ? "," : ".";
  } catch {
    return ".";
  }
}

function normalizeNumericText(value: Primitive): string {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw invalidNumber("Expected finite number");
    }
    return String(value);
  }

  if (typeof value === "boolean") {
    return value ? "1" : "0";
  }

  if (isBlank(value)) {
    throw invalidValue("Expected numeric text");
  }

  let text = String(value).trim().replace(/\u2212/g, "-").replace(/[\s\u00a0\u202f]/g, "");
  const negativeParentheses = /^\(.*\)$/.test(text);
  if (negativeParentheses) {
    text = text.slice(1, -1);
  }

  text = text.replace(/[^0-9+\-.,]/g, "");
  const sign = negativeParentheses || text.startsWith("-") ? "-" : text.startsWith("+") ? "+" : "";
  text = text.replace(/[+\-]/g, "");

  if (!/[0-9]/.test(text)) {
    throw invalidValue("Expected numeric text");
  }

  const separators = [...text].map((char, index) => ({ char, index })).filter((item) => item.char === "." || item.char === ",");
  if (separators.length === 0) {
    return `${sign}${text}`;
  }

  const localeDecimal = runtimeDecimalSeparator();
  const hasComma = text.includes(",");
  const hasDot = text.includes(".");
  let decimalIndex = -1;

  if (hasComma && hasDot) {
    decimalIndex = separators[separators.length - 1].index;
  } else {
    const separator = hasComma ? "," : ".";
    const positions = separators.map((item) => item.index);
    const lastPosition = positions[positions.length - 1];
    const parts = text.split(separator);
    const groupedThousands = parts.length > 1 && parts.slice(1).every((part) => part.length === 3);

    if (positions.length > 1) {
      decimalIndex = groupedThousands ? -1 : lastPosition;
    } else if (separator === localeDecimal) {
      decimalIndex = lastPosition;
    } else {
      decimalIndex = parts[1]?.length === 3 && parts[0].length >= 1 && parts[0].length <= 3 ? -1 : lastPosition;
    }
  }

  if (decimalIndex === -1) {
    return `${sign}${text.replace(/[.,]/g, "")}`;
  }

  const integerPart = text.slice(0, decimalIndex).replace(/[.,]/g, "");
  const decimalPart = text.slice(decimalIndex + 1).replace(/[.,]/g, "");
  return `${sign}${integerPart || "0"}.${decimalPart}`;
}

function parseNumber(value: Primitive): number {
  const normalized = normalizeNumericText(value);
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    throw invalidNumber("Expected numeric text");
  }
  return parsed;
}

function generateNorm(meanValue: number, standardDeviation: number, outlierRate?: Primitive): number {
  if (!Number.isFinite(meanValue) || !Number.isFinite(standardDeviation)) {
    throw invalidValue("Mean and standard deviation must be numeric");
  }
  if (standardDeviation <= 0) {
    throw invalidNumber("Standard deviation must be positive");
  }
  const rate = parseOptionalProbability(outlierRate);
  let value = normalSample(meanValue, standardDeviation);
  if (rate > 0 && Math.random() < rate) {
    value += normalSample(0, standardDeviation * 3);
  }
  return value;
}

function parseArrayCount(count: Primitive): number {
  const parsed = parseOptionalInteger(count, 0);
  if (parsed < 1) {
    throw invalidNumber("Count must be >= 1");
  }
  return parsed;
}

function generateNormArray(count: Primitive, meanValue: number, standardDeviation: number, outlierRate?: Primitive): ExcelOutput {
  const rows = parseArrayCount(count);
  return Array.from({ length: rows }, () => [generateNorm(meanValue, standardDeviation, outlierRate)]);
}

function generateInt(minimum?: Primitive, maximum?: Primitive, outlierRate?: Primitive): number {
  const minValue = parseOptionalInteger(minimum, -2147483648);
  const maxValue = parseOptionalInteger(maximum, 2147483647);
  const rate = parseOptionalProbability(outlierRate);
  if (minValue > maxValue) {
    throw invalidNumber("Minimum cannot exceed maximum");
  }
  let value = Math.floor(Math.random() * (maxValue - minValue + 1)) + minValue;
  if (rate > 0 && Math.random() < rate) {
    const width = Math.max(1, maxValue - minValue);
    const offset = Math.floor(Math.random() * ((2 * width) + 1)) - width;
    value = Math.max(-2147483648, Math.min(2147483647, value + offset));
  }
  return value;
}

function generateIntArray(count: Primitive, minimum: Primitive, maximum: Primitive, outlierRate?: Primitive): ExcelOutput {
  const rows = parseArrayCount(count);
  return Array.from({ length: rows }, () => [generateInt(minimum, maximum, outlierRate)]);
}

function normDistRange(meanValue: number, standardDeviation: number, lowerBound: Primitive, upperBound: Primitive): number {
  if (!Number.isFinite(meanValue) || !Number.isFinite(standardDeviation)) {
    throw invalidValue("Mean and standard deviation must be numeric");
  }
  if (standardDeviation <= 0) {
    throw invalidNumber("Standard deviation must be positive");
  }
  const lower = isBlank(lowerBound) ? Number.NEGATIVE_INFINITY : tryGetNumber(lowerBound);
  const upper = isBlank(upperBound) ? Number.POSITIVE_INFINITY : tryGetNumber(upperBound);
  if (lower === null || upper === null) {
    throw invalidValue("Invalid interval bound");
  }
  if (lower > upper) {
    throw invalidNumber("Lower bound cannot exceed upper bound");
  }
  return jStat.normal.cdf(upper, meanValue, standardDeviation) - jStat.normal.cdf(lower, meanValue, standardDeviation);
}

function averageW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  validateWeights(values, weights);
  return weightedMean(values, weights);
}

function harmMeanW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  validateWeights(values, weights);
  if (values.some((value, index) => weights[index] > 0 && value <= 0)) {
    throw invalidNumber("Weighted harmonic mean requires positive values");
  }
  const sumWeights = weights.reduce((sum, value) => sum + value, 0);
  const denominator = values.reduce((sum, value, index) => sum + (weights[index] === 0 ? 0 : weights[index] / value), 0);
  return sumWeights / denominator;
}

function geoMeanW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  validateWeights(values, weights);
  if (values.some((value, index) => weights[index] > 0 && value <= 0)) {
    throw invalidNumber("Weighted geometric mean requires positive values");
  }
  const sumWeights = weights.reduce((sum, value) => sum + value, 0);
  const logWeightedSum = values.reduce((sum, value, index) => sum + (weights[index] === 0 ? 0 : weights[index] * Math.log(value)), 0);
  return Math.exp(logWeightedSum / sumWeights);
}

function varPW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  return weightedVariance(values, weights, false);
}

function varSW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  return weightedVariance(values, weights, true);
}

function stdevPW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  return Math.sqrt(varPW(valuesInput, weightsInput));
}

function stdevSW(valuesInput: ExcelInput, weightsInput: ExcelInput): number {
  return Math.sqrt(varSW(valuesInput, weightsInput));
}

function varcoef(valuesInput: ExcelInput): number {
  const values = readNumericVector(valuesInput);
  if (values.length < 1) {
    throw invalidValue("At least one observation is required");
  }
  const avg = mean(values);
  if (avg === 0) {
    throw divisionByZero();
  }
  return Math.sqrt(populationVariance(values)) / avg;
}

function varcoefS(valuesInput: ExcelInput): number {
  const values = readNumericVector(valuesInput);
  if (values.length < 2) {
    throw invalidValue("At least two observations are required");
  }
  const avg = mean(values);
  if (avg === 0) {
    throw divisionByZero();
  }
  return Math.sqrt(sampleVariance(values)) / avg;
}

function varcoefW(valuesInput: ExcelInput, weightsInput: ExcelInput, sample = false): number {
  const [values, weights] = readPairedNumericWeights(valuesInput, weightsInput);
  validateWeights(values, weights);
  const avg = weightedMean(values, weights);
  if (avg === 0) {
    throw divisionByZero();
  }
  return Math.sqrt(weightedVariance(values, weights, sample)) / avg;
}

function percentileIncIfs(valuesInput: ExcelInput, quantile: number, ...criteriaArgs: Primitive[]): number {
  if (quantile < 0 || quantile > 1) {
    throw invalidNumber("Quantile must be in [0,1]");
  }
  const filtered = applyFilters(valuesInput, criteriaArgs).sort((a, b) => a - b);
  return percentileInc(filtered, quantile);
}

function percentileExcIfs(valuesInput: ExcelInput, quantile: number, ...criteriaArgs: Primitive[]): number {
  if (quantile <= 0 || quantile >= 1) {
    throw invalidNumber("Quantile must be in (0,1)");
  }
  const filtered = applyFilters(valuesInput, criteriaArgs).sort((a, b) => a - b);
  const rank = quantile * (filtered.length + 1);
  if (rank < 1 || rank > filtered.length) {
    throw invalidNumber("Exclusive percentile rank is out of range");
  }
  return percentileExc(filtered, quantile);
}

function shapiroStatistic(sortedSample: number[]): number {
  const n = sortedSample.length;
  if (n === 3) {
    const range = sortedSample[2] - sortedSample[0];
    const middleDelta = sortedSample[1] - sortedSample[0];
    return range === 0 ? 1 : 0.75 * ((range - (2 * middleDelta)) ** 2) / (range ** 2);
  }

  const half = Math.floor(n / 2);
  const m = new Array<number>(n);
  for (let i = 0; i < n; i += 1) {
    const p = ((i + 1) - 0.375) / (n + 0.25);
    m[i] = jStat.normal.inv(p, 0, 1);
  }

  const mSumSquares = m.reduce((sum, value) => sum + value * value, 0);
  const u = 1 / Math.sqrt(n);
  const weights = new Array<number>(half).fill(0);
  weights[0] = polynomial([0.221157, -0.147981, -2.07119, 4.434685, -2.706056], u);
  if (half > 1) {
    weights[1] = polynomial([0.042981, -0.293762, -1.752461, 5.682633, -3.582633], u);
  }

  const upperExpected = Array.from({ length: half }, (_, index) => m[n - 1 - index]);
  const remainderDenominator = 1 - (2 * weights[0] * weights[0]) - (half > 1 ? 2 * weights[1] * weights[1] : 0);
  const remainderNumerator = mSumSquares - (2 * upperExpected[0] * upperExpected[0]) - (half > 1 ? 2 * upperExpected[1] * upperExpected[1] : 0);
  const scale = Math.sqrt(remainderNumerator / remainderDenominator);

  for (let i = 2; i < half; i += 1) {
    weights[i] = upperExpected[i] / scale;
  }

  let numerator = 0;
  for (let i = 0; i < half; i += 1) {
    numerator += weights[i] * (sortedSample[n - 1 - i] - sortedSample[i]);
  }

  const avg = mean(sortedSample);
  const denominator = sortedSample.reduce((sum, value) => sum + (value - avg) ** 2, 0);
  if (denominator === 0) {
    return 1;
  }

  return Math.max(0, Math.min(1, (numerator * numerator) / denominator));
}

function shapiroPValue(w: number, n: number): number {
  let adjustedW = Math.max(1e-12, Math.min(1 - 1e-12, w));
  if (n === 3) {
    const exact = 1 - ((6 / Math.PI) * Math.acos(Math.sqrt(adjustedW)));
    return Math.max(0, Math.min(1, exact));
  }

  let y = Math.log(1 - adjustedW);
  let mu: number;
  let sigma: number;

  if (n <= 11) {
    const gamma = -2.273 + (0.459 * n);
    if (y >= gamma) {
      return 1e-12;
    }
    y = -Math.log(gamma - y);
    mu = polynomial([0.544, -0.39978, 0.025054, -0.0006714], n);
    sigma = Math.exp(polynomial([1.3822, -0.77857, 0.062767, -0.0020322], n));
  } else {
    const logN = Math.log(n);
    mu = polynomial([-1.5861, -0.31082, -0.083751, 0.0038915], logN);
    sigma = Math.exp(polynomial([-0.4803, -0.082676, 0.0030302], logN));
  }

  const z = (y - mu) / sigma;
  return Math.max(0, Math.min(1, 1 - jStat.normal.cdf(z, 0, 1)));
}

function polynomial(coefficients: number[], x: number): number {
  let value = 0;
  let power = 1;
  for (const coefficient of coefficients) {
    value += coefficient * power;
    power *= x;
  }
  return value;
}

function shapiroWilk(valuesInput: ExcelInput, hasHeader?: Primitive): ExcelOutput {
  const sample = readNumericVector(valuesInput, parseHeaderMode(hasHeader)).sort((a, b) => a - b);
  if (sample.length < 3 || sample.length > 5000) {
    throw invalidValue("Shapiro-Wilk requires between 3 and 5000 observations");
  }
  const w = shapiroStatistic(sample);
  const p = shapiroPValue(w, sample.length);
  return buildRows([["W", w], ["p", p]]);
}

function parseDistribution(input?: Primitive): string {
  if (isBlank(input)) {
    return "normal";
  }
  const code = tryGetNumber(input);
  switch (code) {
    case 0: return "normal";
    case 1: return "lognormal";
    case 2: return "exponential";
    case 3: return "uniform";
    case 4: return "weibull";
    default: throw invalidValue("Unsupported distribution");
  }
}

function ksStatistic(sortedSample: number[], cdf: (value: number) => number): number {
  const n = sortedSample.length;
  let d = 0;
  for (let i = 0; i < n; i += 1) {
    const f = Math.max(0, Math.min(1, cdf(sortedSample[i])));
    const upper = (i + 1) / n;
    const lower = i / n;
    d = Math.max(d, Math.abs(upper - f), Math.abs(f - lower));
  }
  return d;
}

function ksPValue(d: number, n: number, distribution: string): number {
  const lambda = distribution === "normal"
    ? ((Math.sqrt(n) - 0.01 + (0.85 / Math.sqrt(n))) * d)
    : ((Math.sqrt(n) + 0.12 + (0.11 / Math.sqrt(n))) * d);
  return kolmogorovComplementaryCdf(lambda);
}

function kolmogorovSmirnov(valuesInput: ExcelInput, distribution?: Primitive, hasHeader?: Primitive): ExcelOutput {
  const sample = readNumericVector(valuesInput, parseHeaderMode(hasHeader)).sort((a, b) => a - b);
  if (sample.length < 5) {
    throw invalidValue("Kolmogorov-Smirnov requires at least 5 observations");
  }

  const parsedDistribution = parseDistribution(distribution);
  let cdf: (value: number) => number;
  if (parsedDistribution === "normal") {
    const avg = mean(sample);
    const sigma = sampleStandardDeviation(sample);
    if (sigma <= 0) {
      throw invalidNumber("Normal distribution fit requires positive standard deviation");
    }
    cdf = (x) => jStat.normal.cdf(x, avg, sigma);
  } else if (parsedDistribution === "lognormal") {
    if (sample.some((value) => value <= 0)) {
      throw invalidNumber("Lognormal fit requires positive values");
    }
    const logs = sample.map((value) => Math.log(value));
    cdf = (x) => x <= 0 ? 0 : jStat.lognormal.cdf(x, mean(logs), sampleStandardDeviation(logs));
  } else if (parsedDistribution === "exponential") {
    if (sample.some((value) => value < 0)) {
      throw invalidNumber("Exponential fit requires non-negative values");
    }
    const avg = mean(sample);
    if (avg <= 0) {
      throw invalidNumber("Exponential fit requires positive mean");
    }
    const lambda = 1 / avg;
    cdf = (x) => x < 0 ? 0 : 1 - Math.exp(-lambda * x);
  } else if (parsedDistribution === "uniform") {
    const minValue = Math.min(...sample);
    const maxValue = Math.max(...sample);
    if (minValue === maxValue) {
      throw invalidNumber("Uniform fit requires non-constant data");
    }
    cdf = (x) => (x <= minValue ? 0 : x >= maxValue ? 1 : (x - minValue) / (maxValue - minValue));
  } else {
    if (sample.some((value) => value <= 0)) {
      throw invalidNumber("Weibull fit requires positive values");
    }
    const shape = weibullShapeEstimate(sample);
    const scale = weibullScaleEstimate(sample, shape);
    cdf = (x) => x <= 0 ? 0 : 1 - Math.exp(-((x / scale) ** shape));
  }

  const d = ksStatistic(sample, cdf);
  const p = ksPValue(d, sample.length, parsedDistribution);
  return buildRows([["D", d], ["p", p]]);
}

function tTest1S(valuesInput: ExcelInput, mu0: number, direction?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const sample = readNumericVector(valuesInput, parseHeaderMode(hasHeader));
  if (sample.length < 2) {
    throw invalidValue("At least two observations are required");
  }
  const avg = mean(sample);
  const s = sampleStandardDeviation(sample);
  const n = sample.length;
  const df = n - 1;
  const t = (avg - mu0) / (s / Math.sqrt(n));
  return buildRows([
    ["mean", avg],
    ["mu0", mu0],
    ["s", s],
    ["n", n],
    ["alpha", alpha],
    ["t", t],
    ["df", df],
    [criticalLabel("t", parsedDirection), criticalT(alpha, df, parsedDirection)],
    ["p", pValueFromT(t, df, parsedDirection)],
  ]);
}

function propTest1S(valuesInput: ExcelInput, pi0: number, direction?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  if (!(pi0 > 0 && pi0 < 1)) {
    throw invalidNumber("pi0 must be in (0,1)");
  }
  const [successes, count] = readBinaryVector(valuesInput, parseHeaderMode(hasHeader));
  if (count < 1) {
    throw invalidValue("At least one observation is required");
  }
  const pHat = successes / count;
  const z = (pHat - pi0) / Math.sqrt((pi0 * (1 - pi0)) / count);
  return buildRows([
    ["phat", pHat],
    ["pi0", pi0],
    ["x", successes],
    ["n", count],
    ["alpha", alpha],
    ["z", z],
    [criticalLabel("z", parsedDirection), criticalZ(alpha, parsedDirection)],
    ["p", pValueFromZ(z, parsedDirection)],
  ]);
}

function wilcoxonPaired(xInput: ExcelInput, yInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, direction?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const [x, y] = readPairedNumericVectors(xInput, yInput, parseHeaderMode(hasHeader));
  const differences = x.map((value, index) => value - y[index]).filter((value) => Math.abs(value) > 1e-12);
  if (differences.length < 1) {
    throw invalidValue("No non-zero paired differences were found");
  }
  const absDiffs = differences.map((value) => Math.abs(value));
  const ranks = midRank(absDiffs);
  const wPlus = differences.reduce((sum, value, index) => sum + (value > 0 ? ranks[index] : 0), 0);
  const wMinus = differences.reduce((sum, value, index) => sum + (value < 0 ? ranks[index] : 0), 0);
  const w = Math.min(wPlus, wMinus);
  const n = differences.length;
  const meanW = (n * (n + 1)) / 4;
  const tieAdjustment = absDiffs
    .reduce((map, value) => map.set(value, (map.get(value) ?? 0) + 1), new Map<number, number>());
  const varianceW = ((n * (n + 1) * ((2 * n) + 1)) / 24) - Array.from(tieAdjustment.values()).filter((count) => count > 1).reduce((sum, count) => sum + (((count ** 3) - count) / 48), 0);
  if (varianceW <= 0) {
    throw invalidNumber("Wilcoxon variance is not positive");
  }
  const centered = wPlus - meanW;
  let z = centered / Math.sqrt(varianceW);
  if (parsedDirection === "two" && Math.abs(z) > 0) {
    z = Math.sign(z) * Math.max(0, Math.abs(centered) - 0.5) / Math.sqrt(varianceW);
  }
  return buildRows([
    ["n", n],
    ["median_diff", median(differences)],
    ["alpha", alpha],
    ["W+", wPlus],
    ["W-", wMinus],
    ["W", w],
    ["z", z],
    [criticalLabel("z", parsedDirection), criticalZ(alpha, parsedDirection)],
    ["p", pValueFromZ(z, parsedDirection)],
    ["r", Math.abs(z) / Math.sqrt(n)],
  ]);
}

function welchTest2SGCore(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, direction?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const groups = Array.from(readGroups(categoriesInput, valuesInput, parseHeaderMode(hasHeader)).entries()).sort(([a], [b]) => a.localeCompare(b));
  if (groups.length !== 2) {
    throw invalidValue("Welch test requires exactly two groups");
  }
  const [groupAName, groupA] = groups[0];
  const [groupBName, groupB] = groups[1];
  if (groupA.length < 2 || groupB.length < 2) {
    throw invalidValue("Each group must contain at least two values");
  }
  const n1 = groupA.length;
  const n2 = groupB.length;
  const mean1 = mean(groupA);
  const mean2 = mean(groupB);
  const sd1 = sampleStandardDeviation(groupA);
  const sd2 = sampleStandardDeviation(groupB);
  const variancePerN1 = (sd1 * sd1) / n1;
  const variancePerN2 = (sd2 * sd2) / n2;
  const t = (mean1 - mean2) / Math.sqrt(variancePerN1 + variancePerN2);
  const df = ((variancePerN1 + variancePerN2) ** 2) / (((variancePerN1 ** 2) / (n1 - 1)) + ((variancePerN2 ** 2) / (n2 - 1)));
  const pooledVariance = ((((n1 - 1) * sd1 * sd1) + ((n2 - 1) * sd2 * sd2)) / (n1 + n2 - 2));
  const pooledSd = Math.sqrt(pooledVariance);
  const d = pooledSd === 0 ? 0 : (mean1 - mean2) / pooledSd;
  const imbalanceRatio = Math.max(n1, n2) / Math.min(n1, n2);
  const r = imbalanceRatio > 1.5 ? Math.sign(t) * Math.sqrt((t * t) / ((t * t) + df)) : d / Math.sqrt((d * d) + 4);
  return buildRows([
    ["group", "n", "mean", "median", "sd", "min", "max"],
    [groupAName, n1, mean1, median(groupA), sd1, Math.min(...groupA), Math.max(...groupA)],
    [groupBName, n2, mean2, median(groupB), sd2, Math.min(...groupB), Math.max(...groupB)],
    ["", "", "", "", "", "", ""],
    ["α", alpha, "", "", "", "", ""],
    ["t", t, "", "", "", "", ""],
    ["df", df, "", "", "", "", ""],
    [criticalLabel("t", parsedDirection), criticalT(alpha, df, parsedDirection), "", "", "", "", ""],
    ["p", pValueFromT(t, df, parsedDirection), "", "", "", "", ""],
    ["d", d, "", "", "", "", ""],
    ["r", r, "", "", "", "", ""],
  ]);
}

function welchTest2SG(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, direction?: Primitive): ExcelOutput {
  try {
    return welchTest2SGCore(categoriesInput, valuesInput, hasHeader, alpha, direction);
  } catch (error) {
    return rectangularRows([
      ["Evalytics diagnostic", "WELCH.TEST.2S.G failed"],
      ["message", error instanceof Error ? error.message : String(error)],
      ["hasHeader", String(hasHeader ?? "")],
      ["alpha", String(alpha ?? "")],
      ["direction", String(direction ?? "")],
      ["categoryItems", flatten(categoriesInput).length],
      ["valueItems", flatten(valuesInput).length],
      ["firstCategory", String(flatten(categoriesInput)[0] ?? "")],
      ["firstValue", String(flatten(valuesInput)[0] ?? "")],
    ], 7);
  }
}

function mannWhitneyG(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, direction?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const groups = Array.from(readGroups(categoriesInput, valuesInput, parseHeaderMode(hasHeader)).entries()).sort(([a], [b]) => a.localeCompare(b));
  if (groups.length !== 2) {
    throw invalidValue("Mann-Whitney requires exactly two groups");
  }
  const [groupAName, groupA] = groups[0];
  const [groupBName, groupB] = groups[1];
  const n1 = groupA.length;
  const n2 = groupB.length;
  if (n1 < 1 || n2 < 1) {
    throw invalidValue("Each group must contain at least one value");
  }

  const combined = [...groupA.map((value) => ({ group: 0, value })), ...groupB.map((value) => ({ group: 1, value }))];
  const ranks = midRank(combined.map((item) => item.value));
  const rankSumA = combined.reduce((sum, item, index) => sum + (item.group === 0 ? ranks[index] : 0), 0);
  const u1 = rankSumA - ((n1 * (n1 + 1)) / 2);
  const u2 = (n1 * n2) - u1;
  const uStatistic = parsedDirection === "left" ? u1 : parsedDirection === "right" ? u2 : Math.min(u1, u2);
  const meanU = (n1 * n2) / 2;
  const counts = combined.reduce((map, item) => map.set(item.value, (map.get(item.value) ?? 0) + 1), new Map<number, number>());
  const tieCorrection = Array.from(counts.values()).filter((count) => count > 1).reduce((sum, count) => sum + (((count ** 3) - count) / (combined.length * (combined.length - 1))), 0);
  const varianceU = ((n1 * n2) / 12) * (((n1 + n2) + 1) - tieCorrection);
  if (varianceU <= 0) {
    throw invalidNumber("Mann-Whitney variance is not positive");
  }
  const continuity = parsedDirection === "two" ? 0.5 : 0;
  const signedU = parsedDirection === "left" ? u1 : parsedDirection === "right" ? u2 : uStatistic;
  let z = (signedU - meanU + continuity) / Math.sqrt(varianceU);
  if (parsedDirection === "right") {
    z = -z;
  }
  return buildRows([
    ["group", "n", "mean", "median", "sd", "min", "max"],
    [groupAName, n1, mean(groupA), median(groupA), sampleStandardDeviation(groupA), Math.min(...groupA), Math.max(...groupA)],
    [groupBName, n2, mean(groupB), median(groupB), sampleStandardDeviation(groupB), Math.min(...groupB), Math.max(...groupB)],
    ["", "", "", "", "", "", ""],
    ["alpha", alpha, "", "", "", "", ""],
    ["U", uStatistic, "", "", "", "", ""],
    ["U1", u1, "", "", "", "", ""],
    ["U2", u2, "", "", "", "", ""],
    ["z", z, "", "", "", "", ""],
    [criticalLabel("z", parsedDirection), criticalZ(alpha, parsedDirection), "", "", "", "", ""],
    ["p", pValueFromZ(z, parsedDirection), "", "", "", "", ""],
    ["r", Math.abs(z) / Math.sqrt(n1 + n2), "", "", "", "", ""],
  ]);
}

function chisqGof(observedInput: ExcelInput, expectedInput: ExcelInput, categoriesInput?: ExcelInput, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const headerMode = parseHeaderMode(hasHeader);
  const observed = readNumericVector(observedInput, headerMode);
  const expected = readNumericVector(expectedInput, headerMode);
  if (observed.length !== expected.length) {
    throw invalidValue("Observed and expected arrays must have the same length");
  }
  if (observed.length < 2) {
    throw invalidValue("At least two categories are required");
  }
  if (observed.some((value) => value < 0 || Math.abs(value - Math.round(value)) > 1e-9)) {
    throw invalidNumber("Observed values must be non-negative integers");
  }
  if (expected.some((value) => value <= 0)) {
    throw invalidNumber("Expected values must be positive");
  }

  let categories: string[];
  if (categoriesInput === null || categoriesInput === undefined || categoriesInput === "") {
    categories = observed.map((_, index) => String(index + 1));
  } else {
    let raw = flatten(categoriesInput);
    if (headerMode === HEADER_HAS && raw.length > 0) {
      raw = raw.slice(1);
    } else if (headerMode === HEADER_AUTO && raw.length === observed.length + 1) {
      raw = raw.slice(1);
    }
    if (raw.length !== observed.length) {
      throw invalidValue("Category labels length mismatch");
    }
    categories = raw.map((item, index) => (isBlank(item) ? String(index + 1) : String(item)));
  }

  const totalObserved = observed.reduce((sum, value) => sum + value, 0);
  const totalExpected = expected.reduce((sum, value) => sum + value, 0);
  const expectedCounts = Math.abs(totalExpected - 1) < 1e-9 ? expected.map((value) => value * totalObserved) : [...expected];
  const contributions = observed.map((value, index) => ((value - expectedCounts[index]) ** 2) / expectedCounts[index]);
  const chiSquare = contributions.reduce((sum, value) => sum + value, 0);
  const df = observed.length - 1;
  const critical = jStat.chisquare.inv(1 - alpha, df);
  const p = 1 - jStat.chisquare.cdf(chiSquare, df);
  const rows: Primitive[][] = [
    ["chi2", chiSquare, "", ""],
    ["df", df, "", ""],
    ["alpha", alpha, "", ""],
    ["chi2_crit", critical, "", ""],
    ["p", p, "", ""],
    ["", "", "", ""],
    ["category", "O", "E", "(O-E)^2/E"],
  ];
  for (let i = 0; i < observed.length; i += 1) {
    rows.push([categories[i], observed[i], expectedCounts[i], contributions[i]]);
  }
  return buildRows(rows);
}

function parsePostHoc(input?: Primitive): string {
  if (isBlank(input)) {
    return "none";
  }
  const code = tryGetNumber(input);
  switch (code) {
    case 0: return "none";
    case 1: return "tukey";
    case 2: return "bonferroni";
    case 3: return "scheffe";
    case 4: return "games-howell";
    default: throw invalidValue("Invalid post-hoc code");
  }
}

function sigLabel(pValue: number, alpha: number): string {
  return pValue < 0.001 ? "***" : pValue < 0.01 ? "**" : pValue < alpha ? "*" : "ns";
}

function anovaG(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, postHoc?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedPostHoc = parsePostHoc(postHoc);
  const groups = Array.from(readGroups(categoriesInput, valuesInput, parseHeaderMode(hasHeader)).entries()).sort(([a], [b]) => a.localeCompare(b));
  if (groups.length < 3) {
    throw invalidValue("ANOVA requires at least three groups");
  }
  if (groups.some(([, values]) => values.length < 2)) {
    throw invalidValue("Each group must contain at least two observations");
  }

  const totalN = groups.reduce((sum, [, values]) => sum + values.length, 0);
  const grandMean = mean(groups.flatMap(([, values]) => values));
  const k = groups.length;
  const dfBetween = k - 1;
  const dfWithin = totalN - k;
  const dfTotal = totalN - 1;
  const ssBetween = groups.reduce((sum, [, values]) => sum + (values.length * ((mean(values) - grandMean) ** 2)), 0);
  const ssWithin = groups.reduce((sum, [, values]) => {
    const avg = mean(values);
    return sum + values.reduce((inner, value) => inner + ((value - avg) ** 2), 0);
  }, 0);
  const ssTotal = ssBetween + ssWithin;
  const msBetween = ssBetween / dfBetween;
  const msWithin = ssWithin / dfWithin;
  const f = msBetween / msWithin;
  const p = 1 - jStat.centralF.cdf(f, dfBetween, dfWithin);
  const fCrit = jStat.centralF.inv(1 - alpha, dfBetween, dfWithin);

  const transformed = groups.map(([name, values]) => [name, values.map((value) => Math.abs(value - median(values)))] as const);
  const leveneGrandMean = mean(transformed.flatMap(([, values]) => values));
  const leveneSsBetween = transformed.reduce((sum, [, values]) => sum + (values.length * ((mean(values) - leveneGrandMean) ** 2)), 0);
  const leveneSsWithin = transformed.reduce((sum, [, values]) => {
    const avg = mean(values);
    return sum + values.reduce((inner, value) => inner + ((value - avg) ** 2), 0);
  }, 0);
  const leveneF = (leveneSsBetween / dfBetween) / (leveneSsWithin / dfWithin);
  const leveneP = 1 - jStat.centralF.cdf(leveneF, dfBetween, dfWithin);
  const heterogenous = leveneP < alpha;

  const etaSquared = ssTotal === 0 ? 0 : ssBetween / ssTotal;
  let omegaSquared = (ssTotal + msWithin) === 0 ? 0 : (ssBetween - (dfBetween * msWithin)) / (ssTotal + msWithin);
  omegaSquared = Math.max(0, omegaSquared);
  const cohensF = etaSquared >= 1 ? Number.POSITIVE_INFINITY : Math.sqrt(etaSquared / Math.max(1e-12, 1 - etaSquared));

  const rows: Primitive[][] = [
    ["group", "n", "mean", "median", "sd", "min", "max"],
    ...groups.map(([name, values]) => [name, values.length, mean(values), median(values), sampleStandardDeviation(values), Math.min(...values), Math.max(...values)]),
    ["", "", "", "", "", "", ""],
    ["source", "SS", "df", "MS", "F", "p", ""],
    ["between", ssBetween, dfBetween, msBetween, f, p, ""],
    ["within", ssWithin, dfWithin, msWithin, "", "", ""],
    ["total", ssTotal, dfTotal, "", "", "", ""],
    ["alpha", alpha, "", "", "", "", ""],
    ["F_crit", fCrit, "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["levene_F", leveneF, "df1", dfBetween, "df2", dfWithin, ""],
    ["levene_p", leveneP, "heterogenous", heterogenous, "", "", ""],
    ["", "", "", "", "", "", ""],
    ["eta2", etaSquared, "", "", "", "", ""],
    ["omega2", omegaSquared, "", "", "", "", ""],
    ["cohens_f", cohensF, "", "", "", "", ""],
  ];

  if (parsedPostHoc !== "none") {
    rows.push(["", "", "", "", "", "", ""]);
    rows.push(["groupA", "groupB", "mean_diff", "p", "sig", "", ""]);
    const m = (groups.length * (groups.length - 1)) / 2;
    for (let i = 0; i < groups.length; i += 1) {
      for (let j = i + 1; j < groups.length; j += 1) {
        const [nameA, valuesA] = groups[i];
        const [nameB, valuesB] = groups[j];
        const diff = mean(valuesA) - mean(valuesB);
        let pValue = 1;
        if (parsedPostHoc === "games-howell") {
          const varA = (sampleStandardDeviation(valuesA) ** 2) / valuesA.length;
          const varB = (sampleStandardDeviation(valuesB) ** 2) / valuesB.length;
          const t = Math.abs(diff) / Math.sqrt(varA + varB);
          const df = ((varA + varB) ** 2) / (((varA ** 2) / (valuesA.length - 1)) + ((varB ** 2) / (valuesB.length - 1)));
          pValue = 2 * (1 - jStat.studentt.cdf(t, df));
        } else if (parsedPostHoc === "scheffe") {
          const fValue = (diff * diff) / ((groups.length - 1) * msWithin * ((1 / valuesA.length) + (1 / valuesB.length)));
          pValue = 1 - jStat.centralF.cdf(fValue, groups.length - 1, dfWithin);
        } else {
          const se = Math.sqrt(msWithin * ((1 / valuesA.length) + (1 / valuesB.length)));
          const t = Math.abs(diff) / se;
          const rawP = 2 * (1 - jStat.studentt.cdf(t, dfWithin));
          pValue = Math.min(1, rawP * m);
        }
        rows.push([nameA, nameB, diff, pValue, sigLabel(pValue, alpha), "", ""]);
      }
    }
  }

  return buildRows(rows);
}

function correlSpearman(xInput: ExcelInput, yInput: ExcelInput, direction?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const [x, y] = readPairedNumericVectors(xInput, yInput, parseHeaderMode(hasHeader));
  if (x.length < 3) {
    throw invalidValue("Spearman correlation requires at least three paired values");
  }
  const xRanks = midRank(x);
  const yRanks = midRank(y);
  const rho = pearsonCorrelation(xRanks, yRanks);
  const df = x.length - 2;
  const t = Math.abs(1 - Math.abs(rho)) < 1e-12 ? Math.sign(rho) * Number.POSITIVE_INFINITY : (rho * Math.sqrt(df)) / Math.sqrt(1 - (rho * rho));
  return buildRows([
    ["rho", rho],
    ["n", x.length],
    ["alpha", alpha],
    ["t", t],
    ["df", df],
    [criticalLabel("t", parsedDirection), criticalT(alpha, df, parsedDirection)],
    ["p", pValueFromT(t, df, parsedDirection)],
  ]);
}

function pendingFeature(name: string): ExcelOutput {
  throw notAvailable(`${name} is not yet ported to Office.js`);
}

function version(): string {
  return `${ADDIN_BRAND} ${ADDIN_RUNTIME} ${ADDIN_VERSION} build ${ADDIN_BUILD} (EVALYTICS)`;
}

function ping(): number {
  return 1;
}

CustomFunctions.associate("VERSION", version);
CustomFunctions.associate("PING", ping);
CustomFunctions.associate("GENERATE_NORM", generateNorm);
CustomFunctions.associate("GENERATE_NORM_ARRAY", generateNormArray);
CustomFunctions.associate("GENERATE_INT", generateInt);
CustomFunctions.associate("GENERATE_INT_ARRAY", generateIntArray);
CustomFunctions.associate("FILL", fill);
CustomFunctions.associate("PARSE_NUMBER", parseNumber);
CustomFunctions.associate("NORM_DIST_RANGE", normDistRange);
CustomFunctions.associate("AVERAGE_W", averageW);
CustomFunctions.associate("HARMEAN_W", harmMeanW);
CustomFunctions.associate("GEOMEAN_W", geoMeanW);
CustomFunctions.associate("VAR_P_W", varPW);
CustomFunctions.associate("VAR_S_W", varSW);
CustomFunctions.associate("STDEV_P_W", stdevPW);
CustomFunctions.associate("STDEV_S_W", stdevSW);
CustomFunctions.associate("VARCOEF", varcoef);
CustomFunctions.associate("VARCOEF_S", varcoefS);
CustomFunctions.associate("VARCOEF_W", (values: ExcelInput, weights: ExcelInput) => varcoefW(values, weights, false));
CustomFunctions.associate("VARCOEF_S_W", (values: ExcelInput, weights: ExcelInput) => varcoefW(values, weights, true));
CustomFunctions.associate("PERCENTILE_INC_IFS", percentileIncIfs);
CustomFunctions.associate("PERCENTILE_EXC_IFS", percentileExcIfs);
CustomFunctions.associate("SHAPIRO_WILK", shapiroWilk);
CustomFunctions.associate("KOLMOGOROV_SMIRNOV", kolmogorovSmirnov);
CustomFunctions.associate("T_TEST_1S", tTest1S);
CustomFunctions.associate("PROP_TEST_1S", propTest1S);
CustomFunctions.associate("WILCOXON_PAIRED", wilcoxonPaired);
CustomFunctions.associate("WELCH_TEST_2S_G", welchTest2SG);
CustomFunctions.associate("MANN_WHITNEY_G", mannWhitneyG);
CustomFunctions.associate("CHISQ_GOF", chisqGof);
CustomFunctions.associate("ANOVA_G", anovaG);
CustomFunctions.associate("CORREL_SPEARMAN", correlSpearman);
CustomFunctions.associate("ANOVA_RM", () => pendingFeature("ANOVA.RM"));
CustomFunctions.associate("ANCOVA_G", () => pendingFeature("ANCOVA.G"));
CustomFunctions.associate("CONTINGENCY_T", () => pendingFeature("CONTINGENCY.T"));
CustomFunctions.associate("CONTINGENCY_G", () => pendingFeature("CONTINGENCY.G"));
CustomFunctions.associate("CORREL_MATRIX", () => pendingFeature("CORREL.MATRIX"));
CustomFunctions.associate("PIVOT_COUNT", () => pendingFeature("PIVOT.COUNT"));
CustomFunctions.associate("PIVOT_SUM", () => pendingFeature("PIVOT.SUM"));
CustomFunctions.associate("PIVOT_AVERAGE", () => pendingFeature("PIVOT.AVERAGE"));
CustomFunctions.associate("PIVOT_MIN", () => pendingFeature("PIVOT.MIN"));
CustomFunctions.associate("PIVOT_MAX", () => pendingFeature("PIVOT.MAX"));
CustomFunctions.associate("PIVOT_MEDIAN", () => pendingFeature("PIVOT.MEDIAN"));
CustomFunctions.associate("PIVOT_PERCENTILE", () => pendingFeature("PIVOT.PERCENTILE"));
CustomFunctions.associate("PIVOT_STDEV_S", () => pendingFeature("PIVOT.STDEV.S"));
CustomFunctions.associate("PIVOT_STDEV_P", () => pendingFeature("PIVOT.STDEV.P"));
CustomFunctions.associate("PIVOT_VAR_S", () => pendingFeature("PIVOT.VAR.S"));
CustomFunctions.associate("PIVOT_VAR_P", () => pendingFeature("PIVOT.VAR.P"));
CustomFunctions.associate("PIVOT_VARCOEF_S", () => pendingFeature("PIVOT.VARCOEF.S"));
CustomFunctions.associate("PIVOT_VARCOEF_P", () => pendingFeature("PIVOT.VARCOEF.P"));
CustomFunctions.associate("PIVOT_CONF_T", () => pendingFeature("PIVOT.CONF.T"));
CustomFunctions.associate("PIVOT_CONF_NORM", () => pendingFeature("PIVOT.CONF.NORM"));
CustomFunctions.associate("PIVOT_MAD", () => pendingFeature("PIVOT.MAD"));
CustomFunctions.associate("PIVOT_IQR", () => pendingFeature("PIVOT.IQR"));
