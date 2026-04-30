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

function parseNumber(value: Primitive, fallback: Primitive = 0): Primitive {
  try {
    const normalized = normalizeNumericText(value);
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed)) {
      return fallback;
    }
    return parsed;
  } catch {
    return fallback;
  }
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

function asRows(input: ExcelInput): Primitive[][] {
  if (Array.isArray(input)) {
    if (input.length === 0) {
      return [];
    }
    if (Array.isArray(input[0])) {
      return input as Primitive[][];
    }
    return (input as Primitive[]).map((value) => [value]);
  }
  return [[input]];
}

function readPivotDimension(input: ExcelInput, dataLength: number, fallbackHeader: string, totalLabel: string): {
  headers: string[];
  rows: string[][];
} {
  if (!Array.isArray(input) || isBlank(input as Primitive) || input === false) {
    return {
      headers: [fallbackHeader],
      rows: new Array(dataLength).fill(null).map(() => [totalLabel])
    };
  }

  const matrix = asRows(input);
  const hasData = matrix.length > 1 && matrix.slice(1).some((row) => row.some((value) => !isBlank(value)));
  if (!hasData) {
    return {
      headers: [fallbackHeader],
      rows: new Array(dataLength).fill(null).map(() => [totalLabel])
    };
  }

  if (matrix.length - 1 !== dataLength) {
    throw invalidValue("Pivot dimension and values must have the same number of data rows");
  }

  return {
    headers: matrix[0].map((value, index) => isBlank(value) ? `${fallbackHeader}_${index + 1}` : String(value)),
    rows: matrix.slice(1).map((row) => row.map((value) => String(value ?? "")))
  };
}

function cellKey(values: Primitive[]): string {
  return JSON.stringify(values.map((value) => String(value ?? "")));
}

function parsePivotDirection(input?: Primitive): Direction {
  if (isBlank(input)) {
    return "two";
  }
  const code = tryGetNumber(input);
  if (code === 0) {
    return "two";
  }
  if (code === -1) {
    return "left";
  }
  if (code === 1) {
    return "right";
  }
  throw invalidValue("Invalid direction");
}

function readPivotData(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): {
  rowHeaders: string[];
  columnHeaders: string[];
  records: { row: string[]; column: string[]; value: Primitive }[];
} {
  const valueMatrix = asRows(valuesInput);
  if (valueMatrix.length < 2) {
    throw invalidValue("Pivot values must include a header and at least one data row");
  }

  const dataLength = valueMatrix.length - 1;
  const rowDimension = readPivotDimension(rowsInput, dataLength, "row", "TOTAL");
  const columnDimension = readPivotDimension(columnsInput, dataLength, "column", "TOTAL");

  const records = [];
  for (let index = 0; index < dataLength; index += 1) {
    const row = rowDimension.rows[index];
    const column = columnDimension.rows[index];
    records.push({ row, column, value: valueMatrix[index + 1]?.[0] });
  }

  return { rowHeaders: rowDimension.headers, columnHeaders: columnDimension.headers, records };
}

function pivotNumericValues(values: Primitive[]): number[] {
  const result: number[] = [];
  for (const value of values) {
    if (isBlank(value)) {
      continue;
    }
    const parsed = tryGetNumber(value);
    if (parsed === null) {
      throw invalidValue("Expected numeric pivot values");
    }
    result.push(parsed);
  }
  return result;
}

function aggregatePivotValues(values: Primitive[], metric: string, quantile?: number, alpha?: number, direction?: Direction): Primitive {
  if (metric === "count") {
    return values.filter((value) => !isBlank(value)).length;
  }

  const numeric = pivotNumericValues(values).sort((a, b) => a - b);
  if (numeric.length === 0) {
    return "";
  }

  switch (metric) {
    case "sum": return numeric.reduce((sum, value) => sum + value, 0);
    case "average": return mean(numeric);
    case "min": return numeric[0];
    case "max": return numeric[numeric.length - 1];
    case "median": return median(numeric);
    case "percentile":
      if (quantile === undefined || quantile <= 0 || quantile >= 1) {
        throw invalidNumber("Quantile must be in (0,1)");
      }
      return percentileInc(numeric, quantile);
    case "stdev.s": return numeric.length < 2 ? "" : sampleStandardDeviation(numeric);
    case "stdev.p": return Math.sqrt(populationVariance(numeric));
    case "var.s": return numeric.length < 2 ? "" : sampleVariance(numeric);
    case "var.p": return populationVariance(numeric);
    case "varcoef.s": {
      if (numeric.length < 2) {
        return "";
      }
      const avg = mean(numeric);
      if (avg === 0) {
        throw divisionByZero();
      }
      return sampleStandardDeviation(numeric) / avg;
    }
    case "varcoef.p": {
      const avg = mean(numeric);
      if (avg === 0) {
        throw divisionByZero();
      }
      return Math.sqrt(populationVariance(numeric)) / avg;
    }
    case "conf.t": {
      if (numeric.length < 2 || alpha === undefined || direction === undefined) {
        return "";
      }
      validateAlpha(alpha);
      return criticalT(alpha, numeric.length - 1, direction) * sampleStandardDeviation(numeric) / Math.sqrt(numeric.length);
    }
    case "conf.norm": {
      if (numeric.length < 2 || alpha === undefined || direction === undefined) {
        return "";
      }
      validateAlpha(alpha);
      return criticalZ(alpha, direction) * sampleStandardDeviation(numeric) / Math.sqrt(numeric.length);
    }
    case "mad": {
      const med = median(numeric);
      return median(numeric.map((value) => Math.abs(value - med)));
    }
    case "iqr": return percentileInc(numeric, 0.75) - percentileInc(numeric, 0.25);
    default: throw invalidValue("Unsupported pivot metric");
  }
}

function pivotTable(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput, metric: string, options: { quantile?: number; alpha?: number; direction?: Direction } = {}): ExcelOutput {
  const { rowHeaders, columnHeaders, records } = readPivotData(rowsInput, columnsInput, valuesInput);
  const rowKeys = new Map<string, string[]>();
  const columnKeys = new Map<string, string[]>();
  const groups = new Map<string, Primitive[]>();

  const totalColumn = columnHeaders.length === 0 ? ["TOTAL"] : ["TOTAL"];
  columnKeys.set(cellKey(totalColumn), totalColumn);

  for (const record of records) {
    const rowKey = cellKey(record.row);
    const column = columnHeaders.length === 0 ? totalColumn : record.column;
    const columnKey = cellKey(column);
    rowKeys.set(rowKey, record.row);
    columnKeys.set(columnKey, column);
    for (const key of [`${rowKey}\u0000${columnKey}`, `${rowKey}\u0000${cellKey(totalColumn)}`, `${cellKey(["TOTAL"])}\u0000${columnKey}`, `${cellKey(["TOTAL"])}\u0000${cellKey(totalColumn)}`]) {
      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key)!.push(record.value);
    }
  }

  rowKeys.set(cellKey(["TOTAL"]), ["TOTAL", ...new Array(Math.max(0, rowHeaders.length - 1)).fill("")]);
  const sortedRows = Array.from(rowKeys.values()).sort((left, right) => {
    if (left[0] === "TOTAL") return 1;
    if (right[0] === "TOTAL") return -1;
    return left.join("\u0000").localeCompare(right.join("\u0000"));
  });
  const sortedColumns = Array.from(columnKeys.values()).sort((left, right) => {
    if (left[0] === "TOTAL") return 1;
    if (right[0] === "TOTAL") return -1;
    return left.join("\u0000").localeCompare(right.join("\u0000"));
  });

  const headerRows: Primitive[][] = [];
  if (columnHeaders.length > 1) {
    for (let level = 0; level < columnHeaders.length - 1; level += 1) {
      headerRows.push([
        ...new Array(rowHeaders.length).fill(""),
        ...sortedColumns.map((column) => column[0] === "TOTAL" ? "" : column[level]),
      ]);
    }
  }
  headerRows.push([...rowHeaders, ...sortedColumns.map((column) => column[0] === "TOTAL" ? "TOTAL" : column[column.length - 1])]);

  const bodyRows = sortedRows.map((row) => {
    const rowKey = cellKey(row[0] === "TOTAL" ? ["TOTAL"] : row);
    return [
      ...row,
      ...sortedColumns.map((column) => {
        const columnKey = cellKey(column);
        const values = groups.get(`${rowKey}\u0000${columnKey}`) ?? [];
        return aggregatePivotValues(values, metric, options.quantile, options.alpha, options.direction);
      }),
    ];
  });

  return rectangularRows([...headerRows, ...bodyRows]);
}

function pivotCount(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "count");
}

function pivotSum(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "sum");
}

function pivotAverage(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "average");
}

function pivotMin(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "min");
}

function pivotMax(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "max");
}

function pivotMedian(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "median");
}

function pivotPercentile(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput, quantile: number): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "percentile", { quantile });
}

function pivotStdevS(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "stdev.s");
}

function pivotStdevP(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "stdev.p");
}

function pivotVarS(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "var.s");
}

function pivotVarP(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "var.p");
}

function pivotVarcoefS(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "varcoef.s");
}

function pivotVarcoefP(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "varcoef.p");
}

function pivotConfT(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput, alpha: number, direction?: Primitive): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "conf.t", { alpha, direction: parsePivotDirection(direction) });
}

function pivotConfNorm(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput, alpha: number, direction?: Primitive): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "conf.norm", { alpha, direction: parsePivotDirection(direction) });
}

function pivotMad(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "mad");
}

function pivotIqr(rowsInput: ExcelInput, columnsInput: ExcelInput, valuesInput: ExcelInput): ExcelOutput {
  return pivotTable(rowsInput, columnsInput, valuesInput, "iqr");
}

function isIntegerCount(value: number): boolean {
  return value >= 0 && Math.abs(value - Math.round(value)) < 1e-9;
}

function completeNumericMatrix(input: ExcelInput, headerMode: HeaderMode, defaultPrefix: string): { names: string[]; rows: number[][] } {
  const matrix = asRows(input);
  if (matrix.length < 1 || matrix[0].length < 2) {
    throw invalidValue("Expected a matrix with at least two columns");
  }

  const columnCount = matrix[0].length;
  const shouldSkipHeader = headerMode === HEADER_HAS || (
    headerMode === HEADER_AUTO &&
    matrix.length > 1 &&
    matrix[0].every((value) => !isBlank(value) && tryGetNumber(value) === null) &&
    matrix.slice(1).some((row) => row.some((value) => !isBlank(value))) &&
    matrix.slice(1).every((row) => row.every((value) => isBlank(value) || tryGetNumber(value) !== null))
  );
  const startRow = shouldSkipHeader ? 1 : 0;
  const names = new Array(columnCount).fill(null).map((_, index) => {
    if (shouldSkipHeader) {
      const label = matrix[0][index];
      return isBlank(label) ? `${defaultPrefix} ${index + 1}` : String(label);
    }
    return `${defaultPrefix} ${index + 1}`;
  });

  const rows: number[][] = [];
  for (let rowIndex = startRow; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex];
    const values = new Array(columnCount).fill(0);
    let hasAny = false;
    let hasBlank = false;
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const value = row[columnIndex];
      if (isBlank(value)) {
        hasBlank = true;
        continue;
      }
      const parsed = tryGetNumber(value);
      if (parsed === null) {
        throw invalidValue("Expected numeric matrix data");
      }
      hasAny = true;
      values[columnIndex] = parsed;
    }
    if (!hasAny) {
      continue;
    }
    if (!hasBlank) {
      rows.push(values);
    }
  }

  if (rows.length < 2) {
    throw invalidValue("Not enough complete rows");
  }

  return { names, rows };
}

function matrixTranspose(matrix: number[][]): number[][] {
  return matrix[0].map((_, columnIndex) => matrix.map((row) => row[columnIndex]));
}

function matrixMultiply(left: number[][], right: number[][]): number[][] {
  const result = new Array(left.length).fill(null).map(() => new Array(right[0].length).fill(0));
  for (let row = 0; row < left.length; row += 1) {
    for (let column = 0; column < right[0].length; column += 1) {
      for (let inner = 0; inner < right.length; inner += 1) {
        result[row][column] += left[row][inner] * right[inner][column];
      }
    }
  }
  return result;
}

function matrixVectorMultiply(matrix: number[][], vector: number[]): number[] {
  return matrix.map((row) => row.reduce((sum, value, index) => sum + (value * vector[index]), 0));
}

function vectorDot(left: number[], right: number[]): number {
  return left.reduce((sum, value, index) => sum + (value * right[index]), 0);
}

function matrixInverse(input: number[][]): number[][] {
  const size = input.length;
  const augmented = input.map((row, rowIndex) => [
    ...row,
    ...new Array(size).fill(0).map((_, columnIndex) => rowIndex === columnIndex ? 1 : 0)
  ]);

  for (let column = 0; column < size; column += 1) {
    let pivotRow = column;
    for (let row = column + 1; row < size; row += 1) {
      if (Math.abs(augmented[row][column]) > Math.abs(augmented[pivotRow][column])) {
        pivotRow = row;
      }
    }
    if (Math.abs(augmented[pivotRow][column]) < 1e-12) {
      throw invalidNumber("Singular model matrix");
    }
    [augmented[column], augmented[pivotRow]] = [augmented[pivotRow], augmented[column]];

    const pivot = augmented[column][column];
    for (let item = 0; item < size * 2; item += 1) {
      augmented[column][item] /= pivot;
    }
    for (let row = 0; row < size; row += 1) {
      if (row === column) {
        continue;
      }
      const factor = augmented[row][column];
      for (let item = 0; item < size * 2; item += 1) {
        augmented[row][item] -= factor * augmented[column][item];
      }
    }
  }

  return augmented.map((row) => row.slice(size));
}

function fitLinearModel(design: number[][], y: number[]): {
  success: boolean;
  coefficients: number[];
  covariance: number[][];
  sse: number;
  df: number;
  mse: number;
} {
  try {
    const xt = matrixTranspose(design);
    const xtx = matrixMultiply(xt, design);
    const xtxInverse = matrixInverse(xtx);
    const xty = xt.map((row) => vectorDot(row, y));
    const coefficients = matrixVectorMultiply(xtxInverse, xty);
    const fitted = design.map((row) => vectorDot(row, coefficients));
    const residuals = y.map((value, index) => value - fitted[index]);
    const sse = vectorDot(residuals, residuals);
    const df = y.length - design[0].length;
    const mse = df > 0 ? sse / df : Number.NaN;
    const covariance = xtxInverse.map((row) => row.map((value) => value * mse));
    return { success: Number.isFinite(mse), coefficients, covariance, sse, df, mse };
  } catch {
    return { success: false, coefficients: [], covariance: [], sse: Number.NaN, df: 0, mse: Number.NaN };
  }
}

function nestedFTest(reduced: ReturnType<typeof fitLinearModel>, full: ReturnType<typeof fitLinearModel>): {
  ss: number;
  df: number;
  ms: number;
  f: number;
  p: number;
} {
  if (!reduced.success || !full.success || full.df <= 0 || full.mse <= 0) {
    return { ss: Number.NaN, df: 0, ms: Number.NaN, f: Number.NaN, p: Number.NaN };
  }
  const ss = Math.max(0, reduced.sse - full.sse);
  const df = reduced.df - full.df;
  if (df <= 0) {
    return { ss: Number.NaN, df: 0, ms: Number.NaN, f: Number.NaN, p: Number.NaN };
  }
  const ms = ss / df;
  const f = ms / full.mse;
  return { ss, df, ms, f, p: 1 - jStat.centralF.cdf(f, df, full.df) };
}

function sigLabelAlpha(pValue: number, alpha: number): string {
  return pValue < 0.001 ? "***" : pValue < 0.01 ? "**" : pValue < alpha ? "*" : "ns";
}

function anovaRm(valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, postHoc?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedPostHoc = parsePostHoc(postHoc);
  const { names, rows: matrix } = completeNumericMatrix(valuesInput, parseHeaderMode(hasHeader), "condition");
  const subjectCount = matrix.length;
  const conditionCount = names.length;
  if (subjectCount < 2 || conditionCount < 2) {
    throw invalidValue("At least two subjects and two conditions are required");
  }

  const allValues = matrix.flat();
  const grandMean = mean(allValues);
  const conditionMeans = names.map((_, column) => mean(matrix.map((row) => row[column])));
  const subjectMeans = matrix.map((row) => mean(row));
  const ssTotal = allValues.reduce((sum, value) => sum + ((value - grandMean) ** 2), 0);
  const ssConditions = subjectCount * conditionMeans.reduce((sum, value) => sum + ((value - grandMean) ** 2), 0);
  const ssSubjects = conditionCount * subjectMeans.reduce((sum, value) => sum + ((value - grandMean) ** 2), 0);
  const ssError = Math.max(0, ssTotal - ssConditions - ssSubjects);
  const dfConditions = conditionCount - 1;
  const dfSubjects = subjectCount - 1;
  const dfError = dfConditions * dfSubjects;
  const dfTotal = (subjectCount * conditionCount) - 1;
  const msConditions = ssConditions / dfConditions;
  const msSubjects = ssSubjects / dfSubjects;
  const msError = ssError / dfError;
  const f = msError === 0 ? 0 : msConditions / msError;
  const p = 1 - jStat.centralF.cdf(f, dfConditions, dfError);
  const fCrit = jStat.centralF.inv(1 - alpha, dfConditions, dfError);
  const eta2 = ssTotal <= 0 ? "" : ssConditions / ssTotal;
  const eta2p = (ssConditions + ssError) <= 0 ? "" : ssConditions / (ssConditions + ssError);
  const omega2 = (ssTotal + msError) <= 0 ? "" : Math.max(0, (ssConditions - (dfConditions * msError)) / (ssTotal + msError));
  const omega2p = Math.max(0, (dfConditions * (f - 1)) / Math.max(1e-12, (dfConditions * (f - 1)) + subjectCount));

  const rows: Primitive[][] = [
    ["POPISNÉ STATISTIKY", "", "", "", "", "", ""],
    ["condition", "n", "mean", "median", "sd", "min", "max"],
    ...names.map((name, column) => {
      const series = matrix.map((row) => row[column]);
      return [name, subjectCount, mean(series), median(series), sampleStandardDeviation(series), Math.min(...series), Math.max(...series)];
    }),
    ["", "", "", "", "", "", ""],
    ["ANOVA S OPAKOVANÝM MĚŘENÍM", "", "", "", "", "", "", "", "", ""],
    ["source", "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"],
    ["conditions", ssConditions, dfConditions, msConditions, f, p, eta2, eta2p, omega2, omega2p],
    ["subjects", ssSubjects, dfSubjects, msSubjects, "", "", "", "", "", ""],
    ["residual", ssError, dfError, msError, "", "", "", "", "", ""],
    ["total", ssTotal, dfTotal, "", "", "", "", "", "", ""],
    ["α", alpha, "", "", "", "", "", "", "", ""],
    ["Fᶜʳⁱᵗ(1−α)", fCrit, "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["POZNÁMKA", ""],
    ["sphericity", "not tested"],
  ];

  if (parsedPostHoc !== "none") {
    const m = (conditionCount * (conditionCount - 1)) / 2;
    rows.push(["", "", "", "", ""]);
    rows.push([`POST-HOC: ${parsedPostHoc.toUpperCase()}${parsedPostHoc === "bonferroni" ? "" : " (BONFERRONI FALLBACK)"}`, "", "", "", ""]);
    rows.push(["condition A", "condition B", "mean_diff", "p", "sig"]);
    for (let i = 0; i < conditionCount; i += 1) {
      for (let j = i + 1; j < conditionCount; j += 1) {
        const differences = matrix.map((row) => row[i] - row[j]);
        const diff = mean(differences);
        const sdDiff = sampleStandardDeviation(differences);
        const se = sdDiff / Math.sqrt(differences.length);
        const t = se === 0 ? 0 : Math.abs(diff) / se;
        const rawP = 2 * (1 - jStat.studentt.cdf(t, differences.length - 1));
        const pValue = Math.min(1, rawP * m);
        rows.push([names[i], names[j], diff, pValue, sigLabelAlpha(pValue, alpha)]);
      }
    }
  }

  return rectangularRows(rows);
}

function hasContingencyTableHeaders(matrix: Primitive[][]): boolean {
  if (matrix.length < 3 || matrix[0].length < 3 || tryGetNumber(matrix[0][0]) !== null) {
    return false;
  }
  const topHeader = matrix[0].slice(1).some((value) => !isBlank(value) && tryGetNumber(value) === null);
  const leftHeader = matrix.slice(1).some((row) => !isBlank(row[0]) && tryGetNumber(row[0]) === null);
  return topHeader && leftHeader && matrix.slice(1).every((row) => row.slice(1).every((value) => tryGetNumber(value) !== null));
}

function contingencyReport(observed: number[][], rowLabels: string[], columnLabels: string[], alpha: number): ExcelOutput {
  const rowCount = observed.length;
  const columnCount = observed[0].length;
  const rowTotals = observed.map((row) => row.reduce((sum, value) => sum + value, 0));
  const columnTotals = new Array(columnCount).fill(0).map((_, column) => observed.reduce((sum, row) => sum + row[column], 0));
  const total = rowTotals.reduce((sum, value) => sum + value, 0);
  if (total <= 0) {
    throw invalidNumber("Observed counts must sum to a positive value");
  }

  const expected = observed.map((row, rowIndex) => row.map((_, columnIndex) => (rowTotals[rowIndex] * columnTotals[columnIndex]) / total));
  let chiSquare = 0;
  for (let row = 0; row < rowCount; row += 1) {
    for (let column = 0; column < columnCount; column += 1) {
      if (expected[row][column] > 0) {
        chiSquare += ((observed[row][column] - expected[row][column]) ** 2) / expected[row][column];
      }
    }
  }
  const df = (rowCount - 1) * (columnCount - 1);
  const critical = jStat.chisquare.inv(1 - alpha, df);
  const p = 1 - jStat.chisquare.cdf(chiSquare, df);
  const pearsonC = Math.sqrt(chiSquare / (chiSquare + total));
  const cramerV = Math.sqrt(chiSquare / (total * Math.min(rowCount - 1, columnCount - 1)));
  const phi = rowCount === 2 && columnCount === 2 ? Math.sqrt(chiSquare / total) : "";

  const tableHeader = ["", ...columnLabels, "Σ"];
  const rows: Primitive[][] = [
    ["OBSERVED CONTINGENCY TABLE", ...new Array(columnCount + 1).fill("")],
    tableHeader,
    ...observed.map((row, index) => [rowLabels[index], ...row, rowTotals[index]]),
    ["Σ", ...columnTotals, total],
    ["", ...new Array(columnCount + 1).fill("")],
    ["EXPECTED COUNTS", ...new Array(columnCount + 1).fill("")],
    tableHeader,
    ...expected.map((row, index) => [rowLabels[index], ...row, rowTotals[index]]),
    ["Σ", ...columnTotals, total],
    ["", ...new Array(columnCount + 1).fill("")],
    ["TEST SUMMARY", ""],
    ["n", total],
    ["df", df],
    ["α", alpha],
    ["χ²", chiSquare],
    ["χ²ᶜʳⁱᵗ(1−α)", critical],
    ["p", p],
    ["", ""],
    ["ASSOCIATION MEASURES", ""],
    ["Pearson C", pearsonC],
    ["Cramér V", cramerV],
    ["phi", phi],
  ];
  return rectangularRows(rows);
}

function contingencyT(tableInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05): ExcelOutput {
  validateAlpha(alpha);
  const matrix = asRows(tableInput);
  const headerMode = parseHeaderMode(hasHeader);
  const hasHeaders = headerMode === HEADER_HAS || (headerMode === HEADER_AUTO && hasContingencyTableHeaders(matrix));
  const rowOffset = hasHeaders ? 1 : 0;
  const columnOffset = hasHeaders ? 1 : 0;
  const dataRows = matrix.length - rowOffset;
  const dataColumns = (matrix[0]?.length ?? 0) - columnOffset;
  if (dataRows < 2 || dataColumns < 2) {
    throw invalidValue("Contingency table must be at least 2 x 2");
  }

  const observed = new Array(dataRows).fill(null).map((_, row) => new Array(dataColumns).fill(0).map((__, column) => {
    const parsed = tryGetNumber(matrix[row + rowOffset][column + columnOffset]);
    if (parsed === null || !isIntegerCount(parsed)) {
      throw invalidValue("Observed counts must be non-negative integers");
    }
    return parsed;
  }));
  const rowLabels = new Array(dataRows).fill(null).map((_, index) => hasHeaders ? String(matrix[index + 1][0] ?? index + 1) : String(index + 1));
  const columnLabels = new Array(dataColumns).fill(null).map((_, index) => hasHeaders ? String(matrix[0][index + 1] ?? index + 1) : String(index + 1));
  return contingencyReport(observed, rowLabels, columnLabels, alpha);
}

function contingencyG(columnInput: ExcelInput, rowInput: ExcelInput, countsInput?: ExcelInput, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  let columns = flatten(columnInput);
  let rows = flatten(rowInput);
  let counts = countsInput === undefined || isBlank(countsInput as Primitive) ? null : flatten(countsInput);
  if (columns.length !== rows.length || (counts && counts.length !== columns.length)) {
    throw invalidValue("Grouped contingency inputs must have the same length");
  }
  const headerMode = parseHeaderMode(hasHeader);
  if (headerMode === HEADER_HAS || (headerMode === HEADER_AUTO && shouldSkipLeadingHeader(columns, (item) => !isBlank(item)) && shouldSkipLeadingHeader(rows, (item) => !isBlank(item)))) {
    columns = columns.slice(1);
    rows = rows.slice(1);
    counts = counts ? counts.slice(1) : null;
  }

  const rowSet = new Set<string>();
  const columnSet = new Set<string>();
  const pairCounts = new Map<string, number>();
  for (let index = 0; index < columns.length; index += 1) {
    if (isBlank(columns[index]) || isBlank(rows[index]) || (counts && isBlank(counts[index]))) {
      continue;
    }
    const rowLabel = String(rows[index]);
    const columnLabel = String(columns[index]);
    let count = 1;
    if (counts) {
      const parsed = tryGetNumber(counts[index]);
      if (parsed === null || !isIntegerCount(parsed)) {
        throw invalidValue("Counts must be non-negative integers");
      }
      count = parsed;
    }
    rowSet.add(rowLabel);
    columnSet.add(columnLabel);
    const key = `${rowLabel}\u0000${columnLabel}`;
    pairCounts.set(key, (pairCounts.get(key) ?? 0) + count);
  }
  const rowLabels = Array.from(rowSet).sort();
  const columnLabels = Array.from(columnSet).sort();
  if (rowLabels.length < 2 || columnLabels.length < 2) {
    throw invalidValue("At least two row and two column categories are required");
  }
  const observed = rowLabels.map((rowLabel) => columnLabels.map((columnLabel) => pairCounts.get(`${rowLabel}\u0000${columnLabel}`) ?? 0));
  return contingencyReport(observed, rowLabels, columnLabels, alpha);
}

function correlMatrix(dataInput: ExcelInput, method?: Primitive, output?: Primitive, pMinimum?: Primitive, hasHeader?: Primitive): ExcelOutput {
  const parsedMethod = parseOptionalInteger(method, 0);
  const parsedOutput = parseOptionalInteger(output, 0);
  if (parsedMethod < 0 || parsedMethod > 1 || parsedOutput < 0 || parsedOutput > 4) {
    throw invalidValue("Invalid correlation matrix option");
  }
  let parsedPMinimum: number | null = null;
  if (!isBlank(pMinimum)) {
    const p = tryGetNumber(pMinimum);
    if (p === null || p < 0 || p > 1) {
      throw invalidValue("Invalid p minimum");
    }
    parsedPMinimum = p;
  }

  const { names, rows: matrix } = completeNumericMatrix(dataInput, parseHeaderMode(hasHeader), "variable");
  if (names.length < 2 || matrix.length < 3) {
    throw invalidValue("Correlation matrix requires at least two variables and three rows");
  }

  const n = matrix.length;
  const valuesByColumn = names.map((_, column) => matrix.map((row) => row[column]));
  if (valuesByColumn.some((values) => values.every((value) => Math.abs(value - values[0]) < 1e-12))) {
    throw invalidNumber("Correlation matrix variables must vary");
  }
  const coefficients = names.map(() => new Array(names.length).fill(0));
  const pValues = names.map(() => new Array(names.length).fill(0));
  const significance = names.map(() => new Array(names.length).fill(""));
  for (let i = 0; i < names.length; i += 1) {
    for (let j = i; j < names.length; j += 1) {
      const coefficient = i === j ? 1 : parsedMethod === 0
        ? pearsonCorrelation(valuesByColumn[i], valuesByColumn[j])
        : pearsonCorrelation(midRank(valuesByColumn[i]), midRank(valuesByColumn[j]));
      const pValue = Math.abs(1 - Math.abs(coefficient)) < 1e-12 ? 0 : 2 * (1 - jStat.studentt.cdf(Math.abs(coefficient * Math.sqrt(n - 2) / Math.sqrt(1 - (coefficient * coefficient))), n - 2));
      coefficients[i][j] = coefficients[j][i] = coefficient;
      pValues[i][j] = pValues[j][i] = pValue;
      significance[i][j] = significance[j][i] = pValue < 0.001 ? "***" : pValue < 0.01 ? "**" : pValue < 0.05 ? "*" : "";
    }
  }
  const visible = (i: number, j: number): boolean => parsedPMinimum === null ? true : i !== j && pValues[i][j] < parsedPMinimum;

  if (parsedOutput === 2 || parsedOutput === 3) {
    const blockHeight = parsedOutput === 3 ? 3 : 2;
    const rows: Primitive[][] = [["", "", ...names]];
    for (let row = 0; row < names.length; row += 1) {
      rows.push([names[row], "r", ...names.map((_, column) => visible(row, column) ? coefficients[row][column] : "")]);
      rows.push(["", "p", ...names.map((_, column) => visible(row, column) ? pValues[row][column] : "")]);
      if (blockHeight === 3) {
        rows.push(["", "sig.", ...names.map((_, column) => visible(row, column) ? significance[row][column] : "")]);
      }
    }
    return rows;
  }

  return [
    ["", ...names],
    ...names.map((name, row) => [
      name,
      ...names.map((_, column) => {
        if (!visible(row, column)) {
          return "";
        }
        if (parsedOutput === 1) {
          return pValues[row][column];
        }
        if (parsedOutput === 4) {
          return `${Number(coefficients[row][column].toFixed(5))}${significance[row][column]}`;
        }
        return coefficients[row][column];
      })
    ])
  ];
}

type AncovaData = {
  groups: string[];
  y: number[];
  covariates: number[][];
  covariateNames: string[];
  groupLabels: string[];
};

function readAncovaData(groupsInput: ExcelInput, yInput: ExcelInput, covariatesInput: ExcelInput, headerMode: HeaderMode): AncovaData {
  let rawGroups = flatten(groupsInput);
  let rawY = flatten(yInput);
  const covariateMatrix = asRows(covariatesInput);
  if (rawGroups.length !== rawY.length || rawGroups.length !== covariateMatrix.length || covariateMatrix[0].length < 1) {
    throw invalidValue("ANCOVA inputs must have matching lengths");
  }
  const shouldSkipHeader = headerMode === HEADER_HAS || (
    headerMode === HEADER_AUTO &&
    shouldSkipLeadingHeader(rawGroups, (item) => !isBlank(item)) &&
    shouldSkipLeadingHeader(rawY, (item) => tryGetNumber(item) !== null) &&
    covariateMatrix[0].every((_, column) => shouldSkipLeadingHeader(covariateMatrix.map((row) => row[column]), (item) => tryGetNumber(item) !== null))
  );
  const startRow = shouldSkipHeader ? 1 : 0;
  const covariateNames = covariateMatrix[0].map((value, index) => shouldSkipHeader && !isBlank(value) ? String(value) : `covariate ${index + 1}`);
  rawGroups = rawGroups.slice(startRow);
  rawY = rawY.slice(startRow);
  const rawCovariates = covariateMatrix.slice(startRow);

  const groups: string[] = [];
  const y: number[] = [];
  const covariates: number[][] = [];
  for (let row = 0; row < rawGroups.length; row += 1) {
    if (isBlank(rawGroups[row]) || isBlank(rawY[row])) {
      continue;
    }
    const yValue = tryGetNumber(rawY[row]);
    if (yValue === null) {
      throw invalidValue("Expected numeric dependent variable");
    }
    const covariateRow: number[] = [];
    let skip = false;
    for (const cell of rawCovariates[row]) {
      if (isBlank(cell)) {
        skip = true;
        break;
      }
      const parsed = tryGetNumber(cell);
      if (parsed === null) {
        throw invalidValue("Expected numeric covariate");
      }
      covariateRow.push(parsed);
    }
    if (skip) {
      continue;
    }
    groups.push(String(rawGroups[row]));
    y.push(yValue);
    covariates.push(covariateRow);
  }
  const groupLabels = Array.from(new Set(groups)).sort();
  if (groups.length < 3 || groupLabels.length < 2 || groupLabels.some((group) => groups.filter((value) => value === group).length < 2)) {
    throw invalidValue("ANCOVA requires at least two groups with at least two rows each");
  }
  return { groups, y, covariates, covariateNames, groupLabels };
}

function ancovaDesign(data: AncovaData, includeGroups: boolean, covariates: number[], interactionCovariate?: number): number[][] {
  const groupDummies = includeGroups ? data.groupLabels.slice(1) : [];
  return data.y.map((_, rowIndex) => {
    const row = [1];
    for (const group of groupDummies) {
      row.push(data.groups[rowIndex] === group ? 1 : 0);
    }
    for (const covariate of covariates) {
      row.push(data.covariates[rowIndex][covariate]);
    }
    if (interactionCovariate !== undefined) {
      for (const group of groupDummies) {
        row.push((data.groups[rowIndex] === group ? 1 : 0) * data.covariates[rowIndex][interactionCovariate]);
      }
    }
    return row;
  });
}

function buildAncovaEffectRow(name: string, test: ReturnType<typeof nestedFTest>, ssTotal: number, ssError: number, mse: number, sampleSize: number): Primitive[] {
  const eta2 = ssTotal <= 0 ? "" : test.ss / ssTotal;
  const eta2p = (test.ss + ssError) <= 0 ? "" : test.ss / (test.ss + ssError);
  const omega2 = (ssTotal + mse) <= 0 ? "" : Math.max(0, (test.ss - (test.df * mse)) / (ssTotal + mse));
  const omega2p = !Number.isFinite(test.f) ? "" : Math.max(0, (test.df * (test.f - 1)) / ((test.df * (test.f - 1)) + sampleSize));
  return [name, test.ss, test.df, test.ms, test.f, test.p, eta2, eta2p, omega2, omega2p];
}

function ancovaG(groupsInput: ExcelInput, valuesInput: ExcelInput, covariatesInput: ExcelInput, postHoc?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedPostHoc = parsePostHoc(postHoc);
  const data = readAncovaData(groupsInput, valuesInput, covariatesInput, parseHeaderMode(hasHeader));
  const covariateIndexes = data.covariateNames.map((_, index) => index);
  const fullModel = fitLinearModel(ancovaDesign(data, true, covariateIndexes), data.y);
  const covariateOnlyModel = fitLinearModel(ancovaDesign(data, false, covariateIndexes), data.y);
  if (!fullModel.success || fullModel.df <= 0 || !covariateOnlyModel.success) {
    throw invalidNumber("ANCOVA model could not be fitted");
  }
  const factorTest = nestedFTest(covariateOnlyModel, fullModel);
  const covariateTests = covariateIndexes.map((index) => {
    const reduced = fitLinearModel(ancovaDesign(data, true, covariateIndexes.filter((item) => item !== index)), data.y);
    return { name: data.covariateNames[index], test: nestedFTest(reduced, fullModel) };
  });
  const interactionTests = covariateIndexes.map((index) => {
    const interactionModel = fitLinearModel(ancovaDesign(data, true, covariateIndexes, index), data.y);
    return { name: data.covariateNames[index], test: nestedFTest(fullModel, interactionModel) };
  });
  const ssTotal = data.y.reduce((sum, value) => sum + ((value - mean(data.y)) ** 2), 0);
  const modelSsTotal = factorTest.ss + covariateTests.reduce((sum, item) => sum + item.test.ss, 0) + fullModel.sse;
  const descriptiveWidth = Math.max(4, 3 + data.covariateNames.length);
  const rows: Primitive[][] = [
    ["POPISNÉ STATISTIKY", ...new Array(descriptiveWidth - 1).fill("")],
    ["group", "n", "y", ...data.covariateNames],
    ...data.groupLabels.map((group) => {
      const indexes = data.groups.map((value, index) => value === group ? index : -1).filter((index) => index >= 0);
      return [group, indexes.length, mean(indexes.map((index) => data.y[index])), ...covariateIndexes.map((covariate) => mean(indexes.map((index) => data.covariates[index][covariate])))];
    }),
    ["", "", "", ""],
    ["ANCOVA", "", "", "", "", "", "", "", "", ""],
    ["source", "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"],
    buildAncovaEffectRow("factor", factorTest, modelSsTotal, fullModel.sse, fullModel.mse, data.y.length),
    ...covariateTests.map((item) => buildAncovaEffectRow(item.name, item.test, modelSsTotal, fullModel.sse, fullModel.mse, data.y.length)),
    ...interactionTests.map((item) => buildAncovaEffectRow(`group × ${item.name}`, item.test, item.test.ss + fullModel.sse, fullModel.sse, fullModel.mse, data.y.length)),
    ["residual", fullModel.sse, fullModel.df, fullModel.mse, "", "", "", "", "", ""],
    ["total", ssTotal, data.y.length - 1, "", "", "", "", "", "", ""],
    ["α", alpha, "", "", "", "", "", "", "", ""],
    ["Fᶜʳⁱᵗ(1−α)", jStat.centralF.inv(1 - alpha, Math.max(1, factorTest.df), fullModel.df), "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["ADJUSTED MEANS", "", "", "", "", ""],
    ["group", "adjusted_mean", "SE", "CI lower", "CI upper", ""],
  ];

  const covariateMeans = covariateIndexes.map((index) => mean(data.covariates.map((row) => row[index])));
  const adjusted = data.groupLabels.map((group) => {
    const design = [1, ...data.groupLabels.slice(1).map((label) => label === group ? 1 : 0), ...covariateMeans];
    const adjustedMean = vectorDot(design, fullModel.coefficients);
    const variance = vectorDot(design, matrixVectorMultiply(fullModel.covariance, design));
    const se = Math.sqrt(Math.max(0, variance));
    const tCrit = jStat.studentt.inv(1 - (alpha / 2), fullModel.df);
    rows.push([group, adjustedMean, se, adjustedMean - (tCrit * se), adjustedMean + (tCrit * se), ""]);
    return { group, mean: adjustedMean, design };
  });

  if (interactionTests.some((item) => Number.isFinite(item.test.p) && item.test.p < alpha)) {
    rows.push(["", ""]);
    rows.push(["WARNING", ""]);
    rows.push(["slope homogeneity", "violated"]);
  }

  if (parsedPostHoc !== "none") {
    const m = (adjusted.length * (adjusted.length - 1)) / 2;
    rows.push(["", "", "", "", ""]);
    rows.push([`POST-HOC: ${parsedPostHoc.toUpperCase()}${parsedPostHoc === "scheffe" || parsedPostHoc === "bonferroni" ? "" : " (BONFERRONI FALLBACK)"}`, "", "", "", ""]);
    rows.push(["group A", "group B", "mean_diff", "p", "sig"]);
    for (let i = 0; i < adjusted.length; i += 1) {
      for (let j = i + 1; j < adjusted.length; j += 1) {
        const diffVector = adjusted[i].design.map((value, index) => value - adjusted[j].design[index]);
        const diff = adjusted[i].mean - adjusted[j].mean;
        const variance = vectorDot(diffVector, matrixVectorMultiply(fullModel.covariance, diffVector));
        const se = Math.sqrt(Math.max(0, variance));
        const t = se === 0 ? 0 : Math.abs(diff) / se;
        const rawP = 2 * (1 - jStat.studentt.cdf(t, fullModel.df));
        const pValue = parsedPostHoc === "scheffe"
          ? Math.min(1, 1 - jStat.centralF.cdf((t * t) / Math.max(1, adjusted.length - 1), adjusted.length - 1, fullModel.df))
          : Math.min(1, rawP * m);
        rows.push([adjusted[i].group, adjusted[j].group, diff, pValue, sigLabelAlpha(pValue, alpha)]);
      }
    }
  }

  return rectangularRows(rows);
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
CustomFunctions.associate("ANOVA_RM", anovaRm);
CustomFunctions.associate("ANCOVA_G", ancovaG);
CustomFunctions.associate("CONTINGENCY_T", contingencyT);
CustomFunctions.associate("CONTINGENCY_G", contingencyG);
CustomFunctions.associate("CORREL_MATRIX", correlMatrix);
CustomFunctions.associate("PIVOT_COUNT", pivotCount);
CustomFunctions.associate("PIVOT_SUM", pivotSum);
CustomFunctions.associate("PIVOT_AVERAGE", pivotAverage);
CustomFunctions.associate("PIVOT_MIN", pivotMin);
CustomFunctions.associate("PIVOT_MAX", pivotMax);
CustomFunctions.associate("PIVOT_MEDIAN", pivotMedian);
CustomFunctions.associate("PIVOT_PERCENTILE", pivotPercentile);
CustomFunctions.associate("PIVOT_STDEV_S", pivotStdevS);
CustomFunctions.associate("PIVOT_STDEV_P", pivotStdevP);
CustomFunctions.associate("PIVOT_VAR_S", pivotVarS);
CustomFunctions.associate("PIVOT_VAR_P", pivotVarP);
CustomFunctions.associate("PIVOT_VARCOEF_S", pivotVarcoefS);
CustomFunctions.associate("PIVOT_VARCOEF_P", pivotVarcoefP);
CustomFunctions.associate("PIVOT_CONF_T", pivotConfT);
CustomFunctions.associate("PIVOT_CONF_NORM", pivotConfNorm);
CustomFunctions.associate("PIVOT_MAD", pivotMad);
CustomFunctions.associate("PIVOT_IQR", pivotIqr);
