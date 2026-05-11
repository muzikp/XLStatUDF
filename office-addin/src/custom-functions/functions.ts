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
const ADDIN_DISPLAY_VERSION = /^\d+$/.test(ADDIN_BUILD) ? `${ADDIN_VERSION}.${ADDIN_BUILD}` : ADDIN_VERSION;

const HEADER_AUTO: HeaderMode = 0;
const HEADER_HAS: HeaderMode = 1;
const AUTH_STORAGE_KEY = "evalytics.auth.snapshot";
const AUTH_OVERRIDE_ALLOW_ALL = true;

type AuthSnapshot = {
  accessToken?: string;
  apiBaseUrl?: string;
  expiresAt?: number | string;
  entitlements?: {
    allowedFunctions?: string[];
    deniedFunctions?: string[];
  };
};

function invalidValue(message = "Invalid value"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
}

function invalidNumber(message = "Invalid number"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber, message);
}

function notAvailable(message = "Not available"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, message);
}

function unauthorized(): Error {
  return notAvailable("#UNAUTHORIZED");
}

function divisionByZero(message = "Division by zero"): Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.divisionByZero, message);
}

function readAuthSnapshot(): AuthSnapshot | null {
  try {
    if (typeof localStorage === "undefined") {
      return null;
    }
    const raw = localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) {
      return null;
    }
    return JSON.parse(raw) as AuthSnapshot;
  } catch {
    return null;
  }
}

async function readAuthSnapshotShared(): Promise<AuthSnapshot | null> {
  try {
    const runtime = (globalThis as any).OfficeRuntime;
    if (runtime?.storage?.getItem) {
      const raw = await runtime.storage.getItem(AUTH_STORAGE_KEY);
      if (raw) {
        return JSON.parse(raw) as AuthSnapshot;
      }
    }
  } catch {
    // Fallback to localStorage below.
  }
  return readAuthSnapshot();
}

function authExpiresAt(snapshot: AuthSnapshot): number {
  if (typeof snapshot.expiresAt === "number") {
    return snapshot.expiresAt;
  }
  if (typeof snapshot.expiresAt === "string") {
    const parsed = Date.parse(snapshot.expiresAt);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function functionListIncludes(list: string[] | undefined, functionName: string): boolean {
  return Array.isArray(list) && (list.includes("*") || list.includes(functionName));
}

async function remoteAuthCheck(functionName: string, snapshot: AuthSnapshot): Promise<boolean> {
  if (!snapshot.accessToken || !snapshot.apiBaseUrl) {
    return false;
  }
  try {
    const response = await fetch(`${snapshot.apiBaseUrl.replace(/\/$/, "")}/auth/check`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        authorization: `Bearer ${snapshot.accessToken}`
      },
      body: JSON.stringify({ functionName })
    });
    if (!response.ok) {
      return false;
    }
    const payload = await response.json() as { allowed?: boolean };
    return payload.allowed === true;
  } catch {
    return false;
  }
}

function isFunctionAuthorized(functionName: string, snapshot: AuthSnapshot | null): boolean {
  if (!snapshot || authExpiresAt(snapshot) <= Date.now()) {
    return false;
  }

  const allowedFunctions = snapshot.entitlements?.allowedFunctions;
  const deniedFunctions = snapshot.entitlements?.deniedFunctions;
  if (functionListIncludes(deniedFunctions, functionName)) {
    return false;
  }

  return functionListIncludes(allowedFunctions, functionName);
}

async function requireAuthorized(functionName: string): Promise<void> {
  if (AUTH_OVERRIDE_ALLOW_ALL) {
    return;
  }
  const snapshot = await readAuthSnapshotShared();
  if (isFunctionAuthorized(functionName, snapshot)) {
    return;
  }
  if (snapshot && await remoteAuthCheck(functionName, snapshot)) {
    return;
  }
  if (!snapshot || authExpiresAt(snapshot) <= Date.now()) {
    throw unauthorized();
  }
  if (!snapshot.entitlements?.allowedFunctions || !isFunctionAuthorized(functionName, snapshot)) {
    throw unauthorized();
  }
}

function protectedFunction<T extends (...args: any[]) => any>(functionName: string, fn: T): T {
  return (async (...args: Parameters<T>): Promise<ReturnType<T>> => {
    await requireAuthorized(functionName);
    return fn(...args);
  }) as T;
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
    shouldSkipLeadingHeader(rawCategories, (item) => !isBlank(item)) &&
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
  return rows.map((row) => row.map(safeCellValue));
}

function rectangularRows(rows: Primitive[][], width?: number): Primitive[][] {
  const columnCount = width ?? rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
  return rows.map((row) => [...row.map(safeCellValue), ...new Array(Math.max(0, columnCount - row.length)).fill("")]);
}

function safeCellValue(value: Primitive): Primitive {
  return typeof value === "number" && !Number.isFinite(value) ? "" : value;
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

function stackG(dataInput: ExcelInput, hasHeader?: Primitive): ExcelOutput {
  const matrix = asRows(dataInput);
  if (matrix.length < 1 || matrix[0].length < 1) {
    throw invalidValue("Expected a matrix with at least one column");
  }

  const headerMode = parseHeaderMode(hasHeader);
  const hasDetectedHeader = headerMode === HEADER_HAS || (
    headerMode === HEADER_AUTO &&
    matrix.length > 1 &&
    matrix[0].some((value) => !isBlank(value)) &&
    matrix[0].every((value) => isBlank(value) || tryGetNumber(value) === null)
  );
  const startRow = hasDetectedHeader ? 1 : 0;
  const columnCount = matrix[0].length;
  const labels = new Array(columnCount).fill(null).map((_, index) => {
    if (hasDetectedHeader) {
      const label = matrix[0][index];
      return isBlank(label) ? `group ${index + 1}` : String(label);
    }
    return `group ${index + 1}`;
  });

  const rows: Primitive[][] = [];
  for (let column = 0; column < columnCount; column += 1) {
    for (let row = startRow; row < matrix.length; row += 1) {
      const value = matrix[row][column];
      if (isBlank(value)) {
        continue;
      }
      rows.push([labels[column], value]);
    }
  }

  if (rows.length < 2) {
    throw invalidValue("No data rows to stack");
  }

  return rectangularRows(rows, 2);
}

function unstackG(groupsInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive): ExcelOutput {
  let groups = flatten(groupsInput);
  let values = flatten(valuesInput);
  if (groups.length !== values.length) {
    throw invalidValue("Groups and values must have the same length");
  }

  const headerMode = parseHeaderMode(hasHeader);
  const hasDetectedHeader = headerMode === HEADER_HAS || (
    headerMode === HEADER_AUTO &&
    groups.length > 1 &&
    !isBlank(groups[0]) &&
    !isBlank(values[0]) &&
    tryGetNumber(groups[0]) === null &&
    tryGetNumber(values[0]) === null
  );
  if (hasDetectedHeader) {
    groups = groups.slice(1);
    values = values.slice(1);
  }

  const categoryOrder: string[] = [];
  const bucket = new Map<string, Primitive[]>();
  for (let index = 0; index < groups.length; index += 1) {
    if (isBlank(groups[index]) || isBlank(values[index])) {
      continue;
    }
    const key = String(groups[index]);
    if (!bucket.has(key)) {
      bucket.set(key, []);
      categoryOrder.push(key);
    }
    bucket.get(key)!.push(values[index]);
  }

  if (categoryOrder.length < 1) {
    throw invalidValue("No non-empty group/value pairs");
  }

  const maxRows = categoryOrder.reduce((maximum, key) => Math.max(maximum, bucket.get(key)?.length ?? 0), 0);
  const rows: Primitive[][] = [categoryOrder];
  for (let row = 0; row < maxRows; row += 1) {
    rows.push(categoryOrder.map((key) => bucket.get(key)?.[row] ?? ""));
  }

  return rectangularRows(rows, categoryOrder.length);
}

function unstackTable(tableInput: ExcelInput, hasColumnTotal?: Primitive, hasRowTotal?: Primitive): ExcelOutput {
  const matrix = asRows(tableInput);
  if (matrix.length < 3 || (matrix[0]?.length ?? 0) < 3) {
    throw invalidValue("Table must include row/column headers and at least 2x2 data");
  }

  const includeColumnTotal = parseOptionalInteger(hasColumnTotal, 1) !== 0;
  const includeRowTotal = parseOptionalInteger(hasRowTotal, 1) !== 0;
  const rowHeaderName = isBlank(matrix[0][0]) ? "row" : String(matrix[0][0]);
  const columnHeadersRaw = matrix[0].slice(1).map((value, index) => isBlank(value) ? `column ${index + 1}` : String(value));
  const rowHeadersRaw = matrix.slice(1).map((row, index) => isBlank(row[0]) ? `row ${index + 1}` : String(row[0]));

  const dataColumnLimit = includeColumnTotal ? Math.max(0, columnHeadersRaw.length - 1) : columnHeadersRaw.length;
  const dataRowLimit = includeRowTotal ? Math.max(0, rowHeadersRaw.length - 1) : rowHeadersRaw.length;
  if (dataColumnLimit < 2 || dataRowLimit < 2) {
    throw invalidValue("Table without totals must still be at least 2x2");
  }

  const rows: Primitive[][] = [[columnHeadersRaw[0] ?? "column", rowHeaderName, "count"]];
  for (let rowIndex = 0; rowIndex < dataRowLimit; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < dataColumnLimit; columnIndex += 1) {
      const raw = matrix[rowIndex + 1]?.[columnIndex + 1];
      if (isBlank(raw)) {
        continue;
      }
      const count = tryGetNumber(raw);
      if (count === null || !isIntegerCount(count)) {
        throw invalidValue("Counts in the table must be non-negative integers");
      }
      rows.push([columnHeadersRaw[columnIndex], rowHeadersRaw[rowIndex], count]);
    }
  }

  if (rows.length < 2) {
    throw invalidValue("No valid table cells to unstack");
  }
  return rectangularRows(rows, 3);
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

function isBlankDimension(values: Primitive[]): boolean {
  return values.every((value) => isBlank(value));
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
    const value = valueMatrix[index + 1]?.[0];
    if (isBlank(value) && (isBlankDimension(row) || isBlankDimension(column))) {
      continue;
    }
    records.push({ row, column, value });
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

const OUTPUT_LANGUAGE_STORAGE_KEY = "evalytics.language";
type OutputLanguage = "cs" | "en";

const outputLabels: Record<OutputLanguage, Record<string, string>> = {
  cs: {
    descriptiveStats: "POPISNÉ STATISTIKY",
    condition: "podmínka",
    n: "n",
    mean: "průměr",
    median: "medián",
    sd: "sd",
    min: "min",
    max: "max",
    groupDescriptives: "POPISNÉ STATISTIKY SKUPIN",
    anovaTable: "TABULKA ANOVA",
    between: "mezi skupinami",
    within: "uvnitř skupin",
    leveneTest: "LEVENEŮV TEST",
    heterogeneous: "heterogenní",
    effectSize: "VELIKOST EFEKTU",
    anovaRm: "ANOVA S OPAKOVANÝM MĚŘENÍM",
    source: "zdroj",
    conditions: "podmínky",
    subjects: "subjekty",
    residual: "rezidua",
    total: "celkem",
    note: "POZNÁMKA",
    sphericity: "sfericita",
    notTested: "netestována",
    group: "skupina",
    factor: "faktor",
    adjustedMeans: "ADJUSTOVANÉ PRŮMĚRY",
    adjustedMean: "adjustovaný průměr",
    ciLower: "dolní CI",
    ciUpper: "horní CI",
    warning: "VAROVÁNÍ",
    slopeHomogeneity: "homogenita sklonů",
    violated: "porušena",
    groupA: "skupina A",
    groupB: "skupina B",
    meanDiff: "rozdíl průměrů",
    sig: "sig."
  },
  en: {
    descriptiveStats: "DESCRIPTIVE STATISTICS",
    condition: "condition",
    n: "n",
    mean: "mean",
    median: "median",
    sd: "sd",
    min: "min",
    max: "max",
    groupDescriptives: "GROUP DESCRIPTIVES",
    anovaTable: "ANOVA TABLE",
    between: "between",
    within: "within",
    leveneTest: "LEVENE TEST",
    heterogeneous: "heterogeneous",
    effectSize: "EFFECT SIZE",
    anovaRm: "REPEATED-MEASURES ANOVA",
    source: "source",
    conditions: "conditions",
    subjects: "subjects",
    residual: "residual",
    total: "total",
    note: "NOTE",
    sphericity: "sphericity",
    notTested: "not tested",
    group: "group",
    factor: "factor",
    adjustedMeans: "ADJUSTED MEANS",
    adjustedMean: "adjusted_mean",
    ciLower: "CI lower",
    ciUpper: "CI upper",
    warning: "WARNING",
    slopeHomogeneity: "slope homogeneity",
    violated: "violated",
    groupA: "group A",
    groupB: "group B",
    meanDiff: "mean_diff",
    sig: "sig"
  }
};

function normalizeOutputLanguage(value: string | null | undefined): OutputLanguage | null {
  if (!value) {
    return null;
  }
  const lower = value.toLowerCase();
  if (lower.startsWith("cs")) {
    return "cs";
  }
  if (lower.startsWith("en")) {
    return "en";
  }
  return null;
}

function currentOutputLanguage(): OutputLanguage {
  try {
    const stored = typeof localStorage === "undefined" ? null : localStorage.getItem(OUTPUT_LANGUAGE_STORAGE_KEY);
    const storedLanguage = normalizeOutputLanguage(stored);
    if (storedLanguage) {
      return storedLanguage;
    }
  } catch {
    // Some Office custom-functions runtimes may not expose localStorage.
  }

  const navigatorLanguage = normalizeOutputLanguage(
    typeof navigator === "undefined" ? null : navigator.language
  );
  if (navigatorLanguage) {
    return navigatorLanguage;
  }

  try {
    const intlLanguage = normalizeOutputLanguage(
      typeof Intl === "undefined" ? null : Intl.DateTimeFormat().resolvedOptions().locale
    );
    if (intlLanguage) {
      return intlLanguage;
    }
  } catch {
    // Some Office custom-functions runtimes may not expose Intl.
  }

  return "en";
}

function outLabel(key: string): string {
  const language = currentOutputLanguage();
  return outputLabels[language][key] ?? outputLabels.en[key] ?? key;
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
    [outLabel("descriptiveStats"), "", "", "", "", "", ""],
    [outLabel("condition"), outLabel("n"), outLabel("mean"), outLabel("median"), outLabel("sd"), outLabel("min"), outLabel("max")],
    ...names.map((name, column) => {
      const series = matrix.map((row) => row[column]);
      return [name, subjectCount, mean(series), median(series), sampleStandardDeviation(series), Math.min(...series), Math.max(...series)];
    }),
    ["", "", "", "", "", "", ""],
    [outLabel("anovaRm"), "", "", "", "", "", "", "", "", ""],
    [outLabel("source"), "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"],
    [outLabel("conditions"), ssConditions, dfConditions, msConditions, f, p, eta2, eta2p, omega2, omega2p],
    [outLabel("subjects"), ssSubjects, dfSubjects, msSubjects, "", "", "", "", "", ""],
    [outLabel("residual"), ssError, dfError, msError, "", "", "", "", "", ""],
    [outLabel("total"), ssTotal, dfTotal, "", "", "", "", "", "", ""],
    ["α", alpha, "", "", "", "", "", "", "", ""],
    ["Fᶜʳⁱᵗ(1−α)", fCrit, "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    [outLabel("note"), ""],
    [outLabel("sphericity"), outLabel("notTested")],
  ];

  if (parsedPostHoc !== "none") {
    const m = (conditionCount * (conditionCount - 1)) / 2;
    rows.push(["", "", "", "", ""]);
    rows.push([`POST-HOC: ${parsedPostHoc.toUpperCase()}${parsedPostHoc === "bonferroni" ? "" : " (BONFERRONI FALLBACK)"}`, "", "", "", ""]);
    rows.push([`${outLabel("condition")} A`, `${outLabel("condition")} B`, outLabel("meanDiff"), "p", outLabel("sig")]);
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
    ["φ", phi],
  ];
  return rectangularRows(rows);
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
  try {
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
    [outLabel("descriptiveStats"), ...new Array(descriptiveWidth - 1).fill("")],
    [outLabel("group"), outLabel("n"), "y", ...data.covariateNames],
    ...data.groupLabels.map((group) => {
      const indexes = data.groups.map((value, index) => value === group ? index : -1).filter((index) => index >= 0);
      return [group, indexes.length, mean(indexes.map((index) => data.y[index])), ...covariateIndexes.map((covariate) => mean(indexes.map((index) => data.covariates[index][covariate])))];
    }),
    ["", "", "", ""],
    ["ANCOVA", "", "", "", "", "", "", "", "", ""],
    [outLabel("source"), "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"],
    buildAncovaEffectRow(outLabel("factor"), factorTest, modelSsTotal, fullModel.sse, fullModel.mse, data.y.length),
    ...covariateTests.map((item) => buildAncovaEffectRow(item.name, item.test, modelSsTotal, fullModel.sse, fullModel.mse, data.y.length)),
    ...interactionTests.map((item) => buildAncovaEffectRow(`${outLabel("group")} × ${item.name}`, item.test, item.test.ss + fullModel.sse, fullModel.sse, fullModel.mse, data.y.length)),
    [outLabel("residual"), fullModel.sse, fullModel.df, fullModel.mse, "", "", "", "", "", ""],
    [outLabel("total"), ssTotal, data.y.length - 1, "", "", "", "", "", "", ""],
    ["α", alpha, "", "", "", "", "", "", "", ""],
    ["Fᶜʳⁱᵗ(1−α)", jStat.centralF.inv(1 - alpha, Math.max(1, factorTest.df), fullModel.df), "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    [outLabel("adjustedMeans"), "", "", "", "", ""],
    [outLabel("group"), outLabel("adjustedMean"), "SE", outLabel("ciLower"), outLabel("ciUpper"), ""],
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
    rows.push([outLabel("warning"), ""]);
    rows.push([outLabel("slopeHomogeneity"), outLabel("violated")]);
  }

  if (parsedPostHoc !== "none") {
    const m = (adjusted.length * (adjusted.length - 1)) / 2;
    rows.push(["", "", "", "", ""]);
    rows.push([`POST-HOC: ${parsedPostHoc.toUpperCase()}${parsedPostHoc === "scheffe" || parsedPostHoc === "bonferroni" ? "" : " (BONFERRONI FALLBACK)"}`, "", "", "", ""]);
    rows.push([outLabel("groupA"), outLabel("groupB"), outLabel("meanDiff"), "p", outLabel("sig")]);
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
  } catch (error) {
    return rectangularRows([
      ["Evalytics diagnostic", "ANCOVA failed"],
      ["message", error instanceof Error ? error.message : String(error)],
      ["groupsItems", flatten(groupsInput).length],
      ["valueItems", flatten(valuesInput).length],
      ["covariateItems", flatten(covariatesInput).length],
      ["postHoc", String(postHoc ?? "")],
      ["alpha", String(alpha ?? "")],
      ["hasHeader", String(hasHeader ?? "")]
    ], 10);
  }
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
    ["μ₀", mu0],
    ["sₓ", s],
    ["n", n],
    ["α", alpha],
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
    ["p̂", pHat],
    ["π₀", pi0],
    ["x", successes],
    ["n", count],
    ["α", alpha],
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
    ["α", alpha],
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
    ["GROUP DESCRIPTIVES", "", "", "", "", "", ""],
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
      ["Evalytics diagnostic", "WELCH.TEST.2S failed"],
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
    ["α", alpha, "", "", "", "", ""],
    ["U", uStatistic, "", "", "", "", ""],
    ["U1", u1, "", "", "", "", ""],
    ["U2", u2, "", "", "", "", ""],
    ["z", z, "", "", "", "", ""],
    [criticalLabel("z", parsedDirection), criticalZ(alpha, parsedDirection), "", "", "", "", ""],
    ["p", pValueFromZ(z, parsedDirection), "", "", "", "", ""],
    ["r", Math.abs(z) / Math.sqrt(n1 + n2), "", "", "", "", ""],
  ]);
}

function kruskalWallis(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05): ExcelOutput {
  validateAlpha(alpha);
  const groups = Array.from(readGroups(categoriesInput, valuesInput, parseHeaderMode(hasHeader)).entries()).sort(([a], [b]) => a.localeCompare(b));
  if (groups.length < 2) {
    throw invalidValue("Kruskal-Wallis requires at least two groups");
  }
  if (groups.some(([, values]) => values.length < 1)) {
    throw invalidValue("Each group must contain at least one value");
  }

  const combined = groups.flatMap(([, values], groupIndex) => values.map((value) => ({ groupIndex, value })));
  const ranks = midRank(combined.map((item) => item.value));
  const n = combined.length;
  if (n < 3) {
    throw invalidValue("Kruskal-Wallis requires at least three observations");
  }

  const rankSums = new Array(groups.length).fill(0);
  for (let i = 0; i < combined.length; i += 1) {
    rankSums[combined[i].groupIndex] += ranks[i];
  }

  const base = rankSums.reduce((sum, rankSum, index) => sum + ((rankSum * rankSum) / groups[index][1].length), 0);
  let h = ((12 / (n * (n + 1))) * base) - (3 * (n + 1));
  const ties = combined.reduce((map, item) => map.set(item.value, (map.get(item.value) ?? 0) + 1), new Map<number, number>());
  const tieCorrectionDenominator = (n ** 3) - n;
  const tieCorrectionNumerator = Array.from(ties.values()).reduce((sum, count) => sum + ((count ** 3) - count), 0);
  const tieCorrection = tieCorrectionDenominator <= 0 ? 1 : 1 - (tieCorrectionNumerator / tieCorrectionDenominator);
  if (tieCorrection > 1e-12) {
    h /= tieCorrection;
  }

  const df = groups.length - 1;
  const critical = jStat.chisquare.inv(1 - alpha, df);
  const p = 1 - jStat.chisquare.cdf(h, df);
  const epsilonSquared = n <= groups.length ? 0 : Math.max(0, (h - groups.length + 1) / (n - groups.length));

  const rows: Primitive[][] = [
    ["group", "n", "rank_sum", "mean_rank"],
    ...groups.map(([name, values], index) => [name, values.length, rankSums[index], rankSums[index] / values.length]),
    ["", "", "", ""],
    ["H", h, "", ""],
    ["df", df, "", ""],
    ["α", alpha, "", ""],
    ["χ²crit(1-α)", critical, "", ""],
    ["p", p, "", ""],
    ["ε²", epsilonSquared, "", ""]
  ];
  return buildRows(rows);
}

function friedmanAnova(valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05): ExcelOutput {
  validateAlpha(alpha);
  const { names, rows } = completeNumericMatrix(valuesInput, parseHeaderMode(hasHeader), "condition");
  const n = rows.length;
  const k = names.length;
  if (n < 2 || k < 3) {
    throw invalidValue("Friedman ANOVA requires at least two blocks and three conditions");
  }

  const rankSums = new Array(k).fill(0);
  let tieCorrectionSum = 0;
  for (const row of rows) {
    const ranks = midRank(row);
    for (let column = 0; column < k; column += 1) {
      rankSums[column] += ranks[column];
    }
    const ties = row.reduce((map, value) => map.set(value, (map.get(value) ?? 0) + 1), new Map<number, number>());
    tieCorrectionSum += Array.from(ties.values()).reduce((sum, count) => sum + ((count ** 3) - count), 0);
  }

  const sumOfSquares = rankSums.reduce((sum, value) => sum + (value * value), 0);
  let q = ((12 / (n * k * (k + 1))) * sumOfSquares) - (3 * n * (k + 1));
  const correctionDenominator = n * k * ((k ** 2) - 1);
  const correction = correctionDenominator <= 0 ? 1 : 1 - (tieCorrectionSum / correctionDenominator);
  if (correction > 1e-12) {
    q /= correction;
  }
  const df = k - 1;
  const critical = jStat.chisquare.inv(1 - alpha, df);
  const p = 1 - jStat.chisquare.cdf(q, df);
  const kendallsW = (n * ((k ** 2) - 1)) <= 0 ? 0 : q / (n * (k - 1));

  return buildRows([
    ["condition", "n", "rank_sum", "mean_rank"],
    ...names.map((name, column) => [name, n, rankSums[column], rankSums[column] / n]),
    ["", "", "", ""],
    ["Q", q, "", ""],
    ["df", df, "", ""],
    ["α", alpha, "", ""],
    ["χ²crit(1-α)", critical, "", ""],
    ["p", p, "", ""],
    ["Kendall.W", kendallsW, "", ""]
  ]);
}

function jonckheereTerpstra(categoriesInput: ExcelInput, valuesInput: ExcelInput, hasHeader?: Primitive, alpha = 0.05, direction?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedDirection = parseDirection(direction);
  const groups = Array.from(readGroups(categoriesInput, valuesInput, parseHeaderMode(hasHeader)).entries());
  if (groups.length < 3) {
    throw invalidValue("Jonckheere-Terpstra requires at least three ordered groups");
  }
  if (groups.some(([, values]) => values.length < 1)) {
    throw invalidValue("Each group must contain at least one value");
  }

  const orderedGroups = groups;
  const wins = (a: number, b: number): number => (a > b ? 1 : a < b ? 0 : 0.5);
  let j = 0;
  for (let i = 0; i < orderedGroups.length - 1; i += 1) {
    for (let k = i + 1; k < orderedGroups.length; k += 1) {
      for (const left of orderedGroups[i][1]) {
        for (const right of orderedGroups[k][1]) {
          j += wins(right, left);
        }
      }
    }
  }

  const nByGroup = orderedGroups.map(([, values]) => values.length);
  const n = nByGroup.reduce((sum, value) => sum + value, 0);
  const s2 = nByGroup.reduce((sum, ni) => sum + (ni * ni), 0);
  const s3 = nByGroup.reduce((sum, ni) => sum + (ni * ni * ni), 0);
  const meanJ = ((n * n) - s2) / 4;
  let varianceJ = ((n * n * ((2 * n) + 3)) - (s3 * 2) - (s2 * 3)) / 72;

  const allValues = orderedGroups.flatMap(([, values]) => values);
  const ties = allValues.reduce((map, value) => map.set(value, (map.get(value) ?? 0) + 1), new Map<number, number>());
  const tieTerm = Array.from(ties.values()).reduce((sum, count) => sum + ((count ** 3) - count), 0);
  if (n > 1) {
    varianceJ -= (tieTerm * ((n * n) - s2)) / (72 * n * (n - 1));
  }
  if (varianceJ <= 0) {
    throw invalidNumber("Jonckheere-Terpstra variance is not positive");
  }

  let z = (j - meanJ) / Math.sqrt(varianceJ);
  if (parsedDirection === "left") {
    z = -z;
  }

  return buildRows([
    ["group", "n"],
    ...orderedGroups.map(([name, values]) => [name, values.length]),
    ["", ""],
    ["J", j],
    ["E(J)", meanJ],
    ["Var(J)", varianceJ],
    ["z", z],
    ["α", alpha],
    [criticalLabel("z", parsedDirection), criticalZ(alpha, parsedDirection)],
    ["p", pValueFromZ(z, parsedDirection)]
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
    ["χ²", chiSquare, "", ""],
    ["df", df, "", ""],
    ["α", alpha, "", ""],
    ["χ²ᶜʳⁱᵗ(1−α)", critical, "", ""],
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
  if (groups.length < 2) {
    throw invalidValue("ANOVA requires at least two groups");
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
    [outLabel("groupDescriptives"), "", "", "", "", "", ""],
    [outLabel("group"), outLabel("n"), outLabel("mean"), outLabel("median"), outLabel("sd"), outLabel("min"), outLabel("max")],
    ...groups.map(([name, values]) => [name, values.length, mean(values), median(values), sampleStandardDeviation(values), Math.min(...values), Math.max(...values)]),
    ["", "", "", "", "", "", ""],
    [outLabel("anovaTable"), "", "", "", "", "", ""],
    [outLabel("source"), "SS", "df", "MS", "F", "p", ""],
    [outLabel("between"), ssBetween, dfBetween, msBetween, f, p, ""],
    [outLabel("within"), ssWithin, dfWithin, msWithin, "", "", ""],
    [outLabel("total"), ssTotal, dfTotal, "", "", "", ""],
    ["α", alpha, "", "", "", "", ""],
    ["Fᶜʳⁱᵗ(1−α)", fCrit, "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    [outLabel("leveneTest"), "", "", "", "", "", ""],
    ["F", leveneF, "df1", dfBetween, "df2", dfWithin, ""],
    ["p", leveneP, outLabel("heterogeneous"), heterogenous, "", "", ""],
    ["", "", "", "", "", "", ""],
    [outLabel("effectSize"), "", "", "", "", "", ""],
    ["η²", etaSquared, "", "", "", "", ""],
    ["ω²", omegaSquared, "", "", "", "", ""],
    ["Cohen f", cohensF, "", "", "", "", ""],
  ];

  if (parsedPostHoc !== "none") {
    rows.push(["", "", "", "", "", "", ""]);
    rows.push([`POST-HOC (${parsedPostHoc})`, "", "", "", "", "", ""]);
    rows.push([outLabel("groupA"), outLabel("groupB"), "mean(A)-mean(B)", "p", outLabel("sig"), "", ""]);
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
    ["ρ", rho],
    ["n", x.length],
    ["α", alpha],
    ["t", t],
    ["df", df],
    [criticalLabel("t", parsedDirection), criticalT(alpha, df, parsedDirection)],
    ["p", pValueFromT(t, df, parsedDirection)],
  ]);
}

function parseRegressionIntercept(options?: Primitive): boolean {
  if (isBlank(options)) {
    return true;
  }
  const numeric = tryGetNumber(options);
  if (numeric !== null) {
    return Math.abs(numeric) > 1e-12;
  }
  const raw = String(options).trim().toLowerCase();
  const normalized = raw.replace(/\s+/g, "");
  if (["intercept=0", "intercept:false", "intercept=false", "nointercept", "withoutintercept"].includes(normalized)) {
    return false;
  }
  if (["intercept=1", "intercept:true", "intercept=true", "withintercept"].includes(normalized)) {
    return true;
  }
  throw invalidValue("Invalid regression options");
}

function hasNumericHeaderRow(matrix: Primitive[][]): boolean {
  return (
    matrix.length > 1 &&
    matrix[0].every((value) => !isBlank(value) && tryGetNumber(value) === null) &&
    matrix.slice(1).some((row) => row.some((value) => !isBlank(value))) &&
    matrix.slice(1).every((row) => row.every((value) => isBlank(value) || tryGetNumber(value) !== null))
  );
}

function readRegressionData(yInput: ExcelInput, xInput: ExcelInput, headerMode: HeaderMode): {
  y: number[];
  x: number[][];
  predictorNames: string[];
  nSourceRows: number;
  nCompleteRows: number;
  droppedRows: number;
} {
  const yMatrix = asRows(yInput);
  const xMatrix = asRows(xInput);
  const xColumnCount = xMatrix[0]?.length ?? 0;
  if ((yMatrix[0]?.length ?? 0) < 1 || xColumnCount < 1) {
    throw invalidValue("Regression inputs must contain at least one column");
  }

  const hasHeader = headerMode === HEADER_HAS || (headerMode === HEADER_AUTO && hasNumericHeaderRow(xMatrix) && hasNumericHeaderRow(yMatrix));
  const yStart = hasHeader ? 1 : 0;
  const xStart = hasHeader ? 1 : 0;
  const sourceRows = Math.min(yMatrix.length - yStart, xMatrix.length - xStart);
  if (sourceRows < 3) {
    throw invalidValue("Regression requires at least three rows");
  }
  if ((yMatrix.length - yStart) !== (xMatrix.length - xStart)) {
    throw invalidValue("y and x ranges must have the same number of rows");
  }

  const predictorNames = new Array(xColumnCount).fill(null).map((_, index) => {
    if (hasHeader) {
      const label = xMatrix[0][index];
      return isBlank(label) ? `x${index + 1}` : String(label);
    }
    return `x${index + 1}`;
  });

  const y: number[] = [];
  const x: number[][] = [];
  for (let row = 0; row < sourceRows; row += 1) {
    const yValue = yMatrix[row + yStart][0];
    const xRow = xMatrix[row + xStart];
    if (isBlank(yValue) || xRow.some((cell) => isBlank(cell))) {
      continue;
    }
    const parsedY = tryGetNumber(yValue);
    if (parsedY === null) {
      throw invalidValue("Dependent variable must be numeric");
    }
    const parsedX = xRow.map((cell) => {
      const parsed = tryGetNumber(cell);
      if (parsed === null) {
        throw invalidValue("Predictor values must be numeric");
      }
      return parsed;
    });
    y.push(parsedY);
    x.push(parsedX);
  }

  if (y.length < 3) {
    throw invalidValue("Regression requires at least three complete rows");
  }
  return {
    y,
    x,
    predictorNames,
    nSourceRows: sourceRows,
    nCompleteRows: y.length,
    droppedRows: sourceRows - y.length
  };
}

function buildRegressionDesign(x: number[][], includeIntercept: boolean): number[][] {
  if (includeIntercept) {
    return x.map((row) => [1, ...row]);
  }
  return x.map((row) => [...row]);
}

function regression(yInput: ExcelInput, xInput: ExcelInput, options?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const includeIntercept = parseRegressionIntercept(options);
  const data = readRegressionData(yInput, xInput, parseHeaderMode(hasHeader));
  const design = buildRegressionDesign(data.x, includeIntercept);
  if (data.nCompleteRows <= design[0].length) {
    throw invalidValue("Not enough rows for the selected number of predictors");
  }
  const model = fitLinearModel(design, data.y);
  if (!model.success || model.df <= 0) {
    throw invalidNumber("Regression model could not be fitted");
  }

  const fitted = design.map((row) => vectorDot(row, model.coefficients));
  const residuals = data.y.map((value, index) => value - fitted[index]);
  const yMean = mean(data.y);
  const ssTotal = data.y.reduce((sum, value) => sum + ((value - yMean) ** 2), 0);
  const ssModel = Math.max(0, ssTotal - model.sse);
  const dfModel = Math.max(0, design[0].length - (includeIntercept ? 1 : 0));
  const msModel = dfModel > 0 ? ssModel / dfModel : Number.NaN;
  const f = (dfModel > 0 && model.mse > 0) ? msModel / model.mse : Number.NaN;
  const pModel = Number.isFinite(f) ? 1 - jStat.centralF.cdf(f, dfModel, model.df) : Number.NaN;
  const r2 = ssTotal <= 0 ? 0 : Math.max(0, 1 - (model.sse / ssTotal));
  const adjustedR2 = model.df > 0 ? 1 - ((1 - r2) * ((data.nCompleteRows - 1) / model.df)) : Number.NaN;
  const rmse = Math.sqrt(model.mse);
  const n = data.nCompleteRows;
  const k = design[0].length;
  const aic = n * Math.log(Math.max(1e-12, model.sse / n)) + (2 * k);
  const bic = n * Math.log(Math.max(1e-12, model.sse / n)) + (k * Math.log(n));
  const tCrit = jStat.studentt.inv(1 - (alpha / 2), model.df);
  const dwNumerator = residuals.slice(1).reduce((sum, value, index) => sum + ((value - residuals[index]) ** 2), 0);
  const dwDenominator = vectorDot(residuals, residuals);
  const dw = dwDenominator <= 0 ? Number.NaN : dwNumerator / dwDenominator;

  const termNames = includeIntercept ? ["Intercept", ...data.predictorNames] : [...data.predictorNames];
  const rows: Primitive[][] = [
    ["MODEL SUMMARY", ""],
    ["n source", data.nSourceRows],
    ["n complete", data.nCompleteRows],
    ["rows dropped", data.droppedRows],
    ["predictors", data.predictorNames.length],
    ["intercept", includeIntercept ? 1 : 0],
    ["R\u00B2", r2],
    ["Adj. R\u00B2", adjustedR2],
    ["RMSE", rmse],
    ["AIC", aic],
    ["BIC", bic],
    ["", ""],
    ["ANOVA", "", "", "", "", ""],
    ["source", "SS", "df", "MS", "F", "p"],
    ["model", ssModel, dfModel, msModel, f, pModel],
    ["residual", model.sse, model.df, model.mse, "", ""],
    ["total", ssTotal, n - 1, "", "", ""],
    ["", ""],
    ["COEFFICIENTS", "", "", "", "", "", ""],
    ["term", "estimate", "std error", "t", "p", "CI low", "CI high"]
  ];

  for (let index = 0; index < model.coefficients.length; index += 1) {
    const estimate = model.coefficients[index];
    const variance = model.covariance[index]?.[index] ?? Number.NaN;
    const se = variance >= 0 ? Math.sqrt(variance) : Number.NaN;
    const tValue = (Number.isFinite(se) && se > 0) ? estimate / se : Number.NaN;
    const pValue = Number.isFinite(tValue) ? 2 * (1 - jStat.studentt.cdf(Math.abs(tValue), model.df)) : Number.NaN;
    rows.push([
      termNames[index] ?? `x${index + 1}`,
      estimate,
      se,
      tValue,
      pValue,
      Number.isFinite(se) ? estimate - (tCrit * se) : "",
      Number.isFinite(se) ? estimate + (tCrit * se) : ""
    ]);
  }

  rows.push(["", ""]);
  rows.push(["DIAGNOSTICS", ""]);
  rows.push(["Durbin-Watson", dw]);
  return rectangularRows(rows);
}

function regressionPredict(yInput: ExcelInput, xInput: ExcelInput, xNewInput: ExcelInput, options?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const includeIntercept = parseRegressionIntercept(options);
  const data = readRegressionData(yInput, xInput, parseHeaderMode(hasHeader));
  const design = buildRegressionDesign(data.x, includeIntercept);
  if (data.nCompleteRows <= design[0].length) {
    throw invalidValue("Not enough rows for the selected number of predictors");
  }
  const model = fitLinearModel(design, data.y);
  if (!model.success || model.df <= 0) {
    throw invalidNumber("Regression model could not be fitted");
  }

  const xNewRowsRaw = asRows(xNewInput);
  const headerMode = parseHeaderMode(hasHeader);
  const xNewHasHeader = headerMode === HEADER_HAS || (headerMode === HEADER_AUTO && hasNumericHeaderRow(xNewRowsRaw));
  const xNewRows = xNewRowsRaw.slice(xNewHasHeader ? 1 : 0).filter((row) => row.some((cell) => !isBlank(cell)));
  if (xNewRows.length < 1) {
    throw invalidValue("xNew range is empty");
  }

  const predictorCount = data.predictorNames.length;
  const isVectorLike = xNewRowsRaw.length === 1 || (xNewRowsRaw[0]?.length ?? 0) === 1;
  let cleaned: number[][] = [];
  if (isVectorLike) {
    const sequence = flatten(xNewInput).filter((value) => !isBlank(value));
    const parsed = sequence.map((value) => {
      const number = tryGetNumber(value);
      if (number === null) {
        throw invalidValue("xNew must contain only numeric values");
      }
      return number;
    });
    if (parsed.length < predictorCount || parsed.length % predictorCount !== 0) {
      throw invalidValue("xNew sequence length must be a multiple of predictor count");
    }
    for (let index = 0; index < parsed.length; index += predictorCount) {
      cleaned.push(parsed.slice(index, index + predictorCount));
    }
  } else {
    cleaned = xNewRows.map((row) => {
      if (row.length < predictorCount) {
        throw invalidValue("xNew row has fewer predictors than the fitted model");
      }
      return data.predictorNames.map((_, index) => {
        const parsed = tryGetNumber(row[index]);
        if (parsed === null) {
          throw invalidValue("xNew must contain only numeric values");
        }
        return parsed;
      });
    });
  }

  const newDesign = buildRegressionDesign(cleaned, includeIntercept);
  const tCrit = jStat.studentt.inv(1 - (alpha / 2), model.df);
  const rows: Primitive[][] = [["row", "prediction", "CI low", "CI high"]];
  for (let index = 0; index < newDesign.length; index += 1) {
    const designRow = newDesign[index];
    const prediction = vectorDot(designRow, model.coefficients);
    const variance = vectorDot(designRow, matrixVectorMultiply(model.covariance, designRow));
    const seMean = Math.sqrt(Math.max(0, variance));
    rows.push([index + 1, prediction, prediction - (tCrit * seMean), prediction + (tCrit * seMean)]);
  }
  return rectangularRows(rows);
}

type RegressionCriterion = "adjr2" | "aic" | "bic";
type RegressionSelectionMethod = "forward" | "backward" | "stepwise";

function parseRegressionSelectionMethod(method?: Primitive): RegressionSelectionMethod {
  if (isBlank(method)) {
    return "forward";
  }
  const numeric = tryGetNumber(method);
  if (numeric !== null) {
    if (numeric === 0) return "forward";
    if (numeric === 1) return "backward";
    if (numeric === 2) return "stepwise";
  }
  const normalized = String(method).trim().toLowerCase();
  if (["forward", "fwd"].includes(normalized)) return "forward";
  if (["backward", "bwd"].includes(normalized)) return "backward";
  if (["stepwise", "step"].includes(normalized)) return "stepwise";
  throw invalidValue("Invalid regression selection method");
}

function parseRegressionCriterion(criterion?: Primitive): RegressionCriterion {
  if (isBlank(criterion)) {
    return "adjr2";
  }
  const numeric = tryGetNumber(criterion);
  if (numeric !== null) {
    if (numeric === 0) return "adjr2";
    if (numeric === 1) return "aic";
    if (numeric === 2) return "bic";
  }
  const normalized = String(criterion).trim().toLowerCase();
  if (["adjr2", "adj.r2", "adjustedr2"].includes(normalized)) return "adjr2";
  if (normalized === "aic") return "aic";
  if (normalized === "bic") return "bic";
  throw invalidValue("Invalid regression selection criterion");
}

function regressionModelSummary(y: number[], x: number[][], predictorNames: string[], selected: number[], includeIntercept: boolean): {
  success: boolean;
  selected: number[];
  names: string[];
  model: ReturnType<typeof fitLinearModel>;
  n: number;
  sst: number;
  sse: number;
  r2: number;
  adjustedR2: number;
  rmse: number;
  aic: number;
  bic: number;
} {
  const n = y.length;
  const selectedX = x.map((row) => selected.map((index) => row[index]));
  const design = buildRegressionDesign(selectedX, includeIntercept);
  const model = fitLinearModel(design, y);
  if (!model.success || model.df <= 0) {
    return {
      success: false,
      selected,
      names: selected.map((index) => predictorNames[index]),
      model,
      n,
      sst: Number.NaN,
      sse: Number.NaN,
      r2: Number.NaN,
      adjustedR2: Number.NaN,
      rmse: Number.NaN,
      aic: Number.NaN,
      bic: Number.NaN
    };
  }
  const yMean = mean(y);
  const sst = y.reduce((sum, value) => sum + ((value - yMean) ** 2), 0);
  const r2 = sst <= 0 ? 0 : Math.max(0, 1 - (model.sse / sst));
  const adjustedR2 = model.df > 0 ? 1 - ((1 - r2) * ((n - 1) / model.df)) : Number.NaN;
  const rmse = Math.sqrt(model.mse);
  const k = design[0].length;
  const aic = n * Math.log(Math.max(1e-12, model.sse / n)) + (2 * k);
  const bic = n * Math.log(Math.max(1e-12, model.sse / n)) + (k * Math.log(n));
  return {
    success: true,
    selected,
    names: selected.map((index) => predictorNames[index]),
    model,
    n,
    sst,
    sse: model.sse,
    r2,
    adjustedR2,
    rmse,
    aic,
    bic
  };
}

function regressionCriterionScore(summary: ReturnType<typeof regressionModelSummary>, criterion: RegressionCriterion): number {
  if (!summary.success) {
    return Number.NEGATIVE_INFINITY;
  }
  if (criterion === "adjr2") {
    return summary.adjustedR2;
  }
  if (criterion === "aic") {
    return -summary.aic;
  }
  return -summary.bic;
}

function regressionSelect(yInput: ExcelInput, xInput: ExcelInput, method?: Primitive, criterion?: Primitive, alpha = 0.05, hasHeader?: Primitive, options?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const includeIntercept = parseRegressionIntercept(options);
  const selectionMethod = parseRegressionSelectionMethod(method);
  const selectionCriterion = parseRegressionCriterion(criterion);
  const data = readRegressionData(yInput, xInput, parseHeaderMode(hasHeader));
  const p = data.predictorNames.length;
  if (p < 2) {
    throw invalidValue("REGRESSION.SELECT requires at least two predictors");
  }

  let selected = selectionMethod === "backward" ? data.predictorNames.map((_, index) => index) : [];
  let best = regressionModelSummary(data.y, data.x, data.predictorNames, selected, includeIntercept);
  const trace: Primitive[][] = [["step", "action", "variable", "criterion", "model_size"]];
  let step = 0;

  const improveByAdd = (): boolean => {
    let candidateBest = best;
    let candidateVar = -1;
    for (let index = 0; index < p; index += 1) {
      if (selected.includes(index)) continue;
      const candidate = regressionModelSummary(data.y, data.x, data.predictorNames, [...selected, index], includeIntercept);
      if (regressionCriterionScore(candidate, selectionCriterion) > regressionCriterionScore(candidateBest, selectionCriterion) + 1e-10) {
        candidateBest = candidate;
        candidateVar = index;
      }
    }
    if (candidateVar >= 0) {
      selected = [...selected, candidateVar];
      best = candidateBest;
      step += 1;
      trace.push([step, "add", data.predictorNames[candidateVar], selectionCriterion === "adjr2" ? best.adjustedR2 : selectionCriterion === "aic" ? best.aic : best.bic, selected.length]);
      return true;
    }
    return false;
  };

  const improveByRemove = (): boolean => {
    if (selected.length <= 1) return false;
    let candidateBest = best;
    let candidateVar = -1;
    for (const current of selected) {
      const next = selected.filter((value) => value !== current);
      if (next.length < 1) continue;
      const candidate = regressionModelSummary(data.y, data.x, data.predictorNames, next, includeIntercept);
      if (regressionCriterionScore(candidate, selectionCriterion) > regressionCriterionScore(candidateBest, selectionCriterion) + 1e-10) {
        candidateBest = candidate;
        candidateVar = current;
      }
    }
    if (candidateVar >= 0) {
      selected = selected.filter((value) => value !== candidateVar);
      best = candidateBest;
      step += 1;
      trace.push([step, "remove", data.predictorNames[candidateVar], selectionCriterion === "adjr2" ? best.adjustedR2 : selectionCriterion === "aic" ? best.aic : best.bic, selected.length]);
      return true;
    }
    return false;
  };

  if (selectionMethod === "forward") {
    while (improveByAdd()) {
      // iterate until no improvement
    }
  } else if (selectionMethod === "backward") {
    while (improveByRemove()) {
      // iterate until no improvement
    }
  } else {
    let changed = true;
    while (changed) {
      changed = false;
      if (improveByAdd()) changed = true;
      if (improveByRemove()) changed = true;
    }
  }

  if (!best.success || best.selected.length < 1) {
    throw invalidNumber("No valid selected model");
  }

  const rows: Primitive[][] = [
    ["SELECTION SUMMARY", ""],
    ["method", selectionMethod],
    ["criterion", selectionCriterion],
    ["predictors total", data.predictorNames.length],
    ["predictors selected", best.selected.length],
    ["selected", best.names.join(", ")],
    ["R²", best.r2],
    ["Adj. R²", best.adjustedR2],
    ["RMSE", best.rmse],
    ["AIC", best.aic],
    ["BIC", best.bic],
    ["", ""],
    ["TRACE", "", "", "", ""],
    ...trace,
    ["", ""],
    ["COEFFICIENTS", "", "", "", "", "", ""],
    ["term", "estimate", "std error", "t", "p", "CI low", "CI high"]
  ];

  const tCrit = jStat.studentt.inv(1 - (alpha / 2), best.model.df);
  const termNames = includeIntercept ? ["Intercept", ...best.names] : [...best.names];
  for (let index = 0; index < best.model.coefficients.length; index += 1) {
    const estimate = best.model.coefficients[index];
    const variance = best.model.covariance[index]?.[index] ?? Number.NaN;
    const se = variance >= 0 ? Math.sqrt(variance) : Number.NaN;
    const tValue = (Number.isFinite(se) && se > 0) ? estimate / se : Number.NaN;
    const pValue = Number.isFinite(tValue) ? 2 * (1 - jStat.studentt.cdf(Math.abs(tValue), best.model.df)) : Number.NaN;
    rows.push([
      termNames[index] ?? `x${index + 1}`,
      estimate,
      se,
      tValue,
      pValue,
      Number.isFinite(se) ? estimate - (tCrit * se) : "",
      Number.isFinite(se) ? estimate + (tCrit * se) : ""
    ]);
  }
  return rectangularRows(rows);
}

type TrendModelType = "linear" | "log" | "exp" | "power";

function parseTrendModelType(modelType?: Primitive): TrendModelType {
  if (isBlank(modelType)) return "linear";
  const numeric = tryGetNumber(modelType);
  if (numeric !== null) {
    if (numeric === 0) return "linear";
    if (numeric === 1) return "log";
    if (numeric === 2) return "exp";
    if (numeric === 3) return "power";
  }
  const normalized = String(modelType).trim().toLowerCase();
  if (["linear", "lin"].includes(normalized)) return "linear";
  if (["log", "logarithmic"].includes(normalized)) return "log";
  if (["exp", "exponential"].includes(normalized)) return "exp";
  if (["power", "pow"].includes(normalized)) return "power";
  throw invalidValue("Invalid trend model type");
}

function readTrendData(yInput: ExcelInput, xInput: ExcelInput, headerMode: HeaderMode): { x: number[]; y: number[] } {
  const [x, y] = readPairedNumericVectors(xInput, yInput, headerMode);
  if (x.length < 3) {
    throw invalidValue("Trend fitting requires at least three paired values");
  }
  return { x, y };
}

function evaluateTrendModel(x: number[], y: number[], modelType: TrendModelType): {
  modelType: TrendModelType;
  equation: string;
  coefficients: Record<string, number>;
  n: number;
  r2: number;
  adjustedR2: number;
  rmse: number;
  aic: number;
  bic: number;
} {
  const n = x.length;
  const transformedX: number[] = [];
  const transformedY: number[] = [];
  for (let i = 0; i < n; i += 1) {
    const xi = x[i];
    const yi = y[i];
    if (modelType === "log" || modelType === "power") {
      if (xi <= 0) throw invalidNumber("Log and power trends require x > 0");
    }
    if (modelType === "exp" || modelType === "power") {
      if (yi <= 0) throw invalidNumber("Exponential and power trends require y > 0");
    }
    transformedX.push(modelType === "log" || modelType === "power" ? Math.log(xi) : xi);
    transformedY.push(modelType === "exp" || modelType === "power" ? Math.log(yi) : yi);
  }

  const design = transformedX.map((value) => [1, value]);
  const fit = fitLinearModel(design, transformedY);
  if (!fit.success || fit.df <= 0) {
    throw invalidNumber("Trend model could not be fitted");
  }
  const intercept = fit.coefficients[0];
  const slope = fit.coefficients[1];

  const predicted = x.map((xi, index) => {
    if (modelType === "linear") return intercept + (slope * xi);
    if (modelType === "log") return intercept + (slope * Math.log(xi));
    if (modelType === "exp") return Math.exp(intercept + (slope * xi));
    return Math.exp(intercept) * (xi ** slope);
  });

  const residuals = y.map((value, index) => value - predicted[index]);
  const sse = vectorDot(residuals, residuals);
  const yMean = mean(y);
  const sst = y.reduce((sum, value) => sum + ((value - yMean) ** 2), 0);
  const r2 = sst <= 0 ? 0 : Math.max(0, 1 - (sse / sst));
  const adjustedR2 = 1 - ((1 - r2) * ((n - 1) / (n - 2)));
  const rmse = Math.sqrt(sse / (n - 2));
  const aic = n * Math.log(Math.max(1e-12, sse / n)) + (2 * 2);
  const bic = n * Math.log(Math.max(1e-12, sse / n)) + (2 * Math.log(n));

  if (modelType === "linear") {
    return {
      modelType,
      equation: `y = ${intercept} + ${slope}*x`,
      coefficients: { intercept, slope },
      n,
      r2,
      adjustedR2,
      rmse,
      aic,
      bic
    };
  }
  if (modelType === "log") {
    return {
      modelType,
      equation: `y = ${intercept} + ${slope}*ln(x)`,
      coefficients: { intercept, slope },
      n,
      r2,
      adjustedR2,
      rmse,
      aic,
      bic
    };
  }
  if (modelType === "exp") {
    return {
      modelType,
      equation: `y = ${Math.exp(intercept)}*exp(${slope}*x)`,
      coefficients: { a: Math.exp(intercept), b: slope, lnA: intercept },
      n,
      r2,
      adjustedR2,
      rmse,
      aic,
      bic
    };
  }
  return {
    modelType,
    equation: `y = ${Math.exp(intercept)}*x^${slope}`,
    coefficients: { a: Math.exp(intercept), b: slope, lnA: intercept },
    n,
    r2,
    adjustedR2,
    rmse,
    aic,
    bic
  };
}

function trendFit(yInput: ExcelInput, xInput: ExcelInput, modelType?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const parsedType = parseTrendModelType(modelType);
  const data = readTrendData(yInput, xInput, parseHeaderMode(hasHeader));
  const fit = evaluateTrendModel(data.x, data.y, parsedType);
  const rows: Primitive[][] = [
    ["TREND FIT", ""],
    ["model", fit.modelType],
    ["equation", fit.equation],
    ["n", fit.n],
    ["R²", fit.r2],
    ["Adj. R²", fit.adjustedR2],
    ["RMSE", fit.rmse],
    ["AIC", fit.aic],
    ["BIC", fit.bic],
    ["", ""],
    ["COEFFICIENTS", ""]
  ];
  for (const [key, value] of Object.entries(fit.coefficients)) {
    rows.push([key, value]);
  }
  return rectangularRows(rows);
}

function parseTrendModelList(modelList?: Primitive): TrendModelType[] {
  if (isBlank(modelList)) {
    return ["linear", "log", "exp", "power"];
  }
  const raw = String(modelList)
    .split(/[;,|]/g)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
  if (raw.length < 1) {
    return ["linear", "log", "exp", "power"];
  }
  const models: TrendModelType[] = [];
  for (const token of raw) {
    models.push(parseTrendModelType(token));
  }
  return Array.from(new Set(models));
}

function trendCompare(yInput: ExcelInput, xInput: ExcelInput, modelList?: Primitive, alpha = 0.05, hasHeader?: Primitive): ExcelOutput {
  validateAlpha(alpha);
  const models = parseTrendModelList(modelList);
  const data = readTrendData(yInput, xInput, parseHeaderMode(hasHeader));
  const rows: Primitive[][] = [["model", "equation", "n", "R²", "Adj. R²", "RMSE", "AIC", "BIC"]];
  for (const model of models) {
    const fit = evaluateTrendModel(data.x, data.y, model);
    rows.push([fit.modelType, fit.equation, fit.n, fit.r2, fit.adjustedR2, fit.rmse, fit.aic, fit.bic]);
  }
  return rectangularRows(rows);
}

function version(): string {
  return `${ADDIN_BRAND} ${ADDIN_RUNTIME} ${ADDIN_DISPLAY_VERSION} (EVALYTICS)`;
}

function ping(): number {
  return 1;
}

CustomFunctions.associate("VERSION", version);
CustomFunctions.associate("PING", ping);
CustomFunctions.associate("GENERATE_NORM", protectedFunction("GENERATE.NORM", generateNorm));
CustomFunctions.associate("GENERATE_NORM_ARRAY", protectedFunction("GENERATE.NORM.ARRAY", generateNormArray));
CustomFunctions.associate("GENERATE_INT", protectedFunction("GENERATE.INT", generateInt));
CustomFunctions.associate("GENERATE_INT_ARRAY", protectedFunction("GENERATE.INT.ARRAY", generateIntArray));
CustomFunctions.associate("FILL", fill);
CustomFunctions.associate("STACK_G", stackG);
CustomFunctions.associate("UNSTACK_G", unstackG);
CustomFunctions.associate("UNSTACK_TABLE", unstackTable);
CustomFunctions.associate("PARSE_NUMBER", parseNumber);
CustomFunctions.associate("NORM_DIST_RANGE", protectedFunction("NORM.DIST.RANGE", normDistRange));
CustomFunctions.associate("AVERAGE_W", protectedFunction("AVERAGE.W", averageW));
CustomFunctions.associate("HARMEAN_W", protectedFunction("HARMEAN.W", harmMeanW));
CustomFunctions.associate("GEOMEAN_W", protectedFunction("GEOMEAN.W", geoMeanW));
CustomFunctions.associate("VAR_P_W", protectedFunction("VAR.P.W", varPW));
CustomFunctions.associate("VAR_S_W", protectedFunction("VAR.S.W", varSW));
CustomFunctions.associate("STDEV_P_W", protectedFunction("STDEV.P.W", stdevPW));
CustomFunctions.associate("STDEV_S_W", protectedFunction("STDEV.S.W", stdevSW));
CustomFunctions.associate("VARCOEF", protectedFunction("VARCOEF", varcoef));
CustomFunctions.associate("VARCOEF_S", protectedFunction("VARCOEF.S", varcoefS));
CustomFunctions.associate("VARCOEF_W", protectedFunction("VARCOEF.W", (values: ExcelInput, weights: ExcelInput) => varcoefW(values, weights, false)));
CustomFunctions.associate("VARCOEF_S_W", protectedFunction("VARCOEF.S.W", (values: ExcelInput, weights: ExcelInput) => varcoefW(values, weights, true)));
CustomFunctions.associate("SHAPIRO_WILK", protectedFunction("SHAPIRO.WILK", shapiroWilk));
CustomFunctions.associate("KOLMOGOROV_SMIRNOV", protectedFunction("KOLMOGOROV.SMIRNOV", kolmogorovSmirnov));
CustomFunctions.associate("T_TEST_1S", protectedFunction("T.TEST.1S", tTest1S));
CustomFunctions.associate("PROP_TEST_1S", protectedFunction("PROP.TEST.1S", propTest1S));
CustomFunctions.associate("WILCOXON_PAIRED", protectedFunction("WILCOXON.PAIRED", wilcoxonPaired));
CustomFunctions.associate("WELCH_TEST_2S_G", protectedFunction("WELCH.TEST.2S", welchTest2SG));
CustomFunctions.associate("MANN_WHITNEY_G", protectedFunction("MANN.WHITNEY", mannWhitneyG));
CustomFunctions.associate("KRUSKAL_WALLIS", protectedFunction("KRUSKAL.WALLIS", kruskalWallis));
CustomFunctions.associate("FRIEDMAN_ANOVA", protectedFunction("FRIEDMAN.ANOVA", friedmanAnova));
CustomFunctions.associate("JONCKHEERE_TERPSTRA", protectedFunction("JONCKHEERE.TERPSTRA", jonckheereTerpstra));
CustomFunctions.associate("CHISQ_GOF", protectedFunction("CHISQ.GOF", chisqGof));
CustomFunctions.associate("ANOVA_G", protectedFunction("ANOVA", anovaG));
CustomFunctions.associate("CORREL_SPEARMAN", protectedFunction("CORREL.SPEARMAN", correlSpearman));
CustomFunctions.associate("REGRESSION", protectedFunction("REGRESSION", regression));
CustomFunctions.associate("REGRESSION_PREDICT", protectedFunction("REGRESSION.PREDICT", regressionPredict));
CustomFunctions.associate("REGRESSION_SELECT", protectedFunction("REGRESSION.SELECT", regressionSelect));
CustomFunctions.associate("TREND_FIT", protectedFunction("TREND.FIT", trendFit));
CustomFunctions.associate("TREND_COMPARE", protectedFunction("TREND.COMPARE", trendCompare));
CustomFunctions.associate("ANOVA_RM", protectedFunction("ANOVA.RM", anovaRm));
CustomFunctions.associate("ANCOVA_G", protectedFunction("ANCOVA", ancovaG));
CustomFunctions.associate("CONTINGENCY_G", protectedFunction("CONTINGENCY", contingencyG));
CustomFunctions.associate("CORREL_MATRIX", protectedFunction("CORREL.MATRIX", correlMatrix));
CustomFunctions.associate("PIVOT_COUNT", protectedFunction("PIVOT.COUNT", pivotCount));
CustomFunctions.associate("PIVOT_SUM", protectedFunction("PIVOT.SUM", pivotSum));
CustomFunctions.associate("PIVOT_AVERAGE", protectedFunction("PIVOT.AVERAGE", pivotAverage));
CustomFunctions.associate("PIVOT_MIN", protectedFunction("PIVOT.MIN", pivotMin));
CustomFunctions.associate("PIVOT_MAX", protectedFunction("PIVOT.MAX", pivotMax));
CustomFunctions.associate("PIVOT_MEDIAN", protectedFunction("PIVOT.MEDIAN", pivotMedian));
CustomFunctions.associate("PIVOT_PERCENTILE", protectedFunction("PIVOT.PERCENTILE", pivotPercentile));
CustomFunctions.associate("PIVOT_STDEV_S", protectedFunction("PIVOT.STDEV.S", pivotStdevS));
CustomFunctions.associate("PIVOT_STDEV_P", protectedFunction("PIVOT.STDEV.P", pivotStdevP));
CustomFunctions.associate("PIVOT_VAR_S", protectedFunction("PIVOT.VAR.S", pivotVarS));
CustomFunctions.associate("PIVOT_VAR_P", protectedFunction("PIVOT.VAR.P", pivotVarP));
CustomFunctions.associate("PIVOT_VARCOEF_S", protectedFunction("PIVOT.VARCOEF.S", pivotVarcoefS));
CustomFunctions.associate("PIVOT_VARCOEF_P", protectedFunction("PIVOT.VARCOEF.P", pivotVarcoefP));
CustomFunctions.associate("PIVOT_CONF_T", protectedFunction("PIVOT.CONF.T", pivotConfT));
CustomFunctions.associate("PIVOT_CONF_NORM", protectedFunction("PIVOT.CONF.NORM", pivotConfNorm));
CustomFunctions.associate("PIVOT_MAD", protectedFunction("PIVOT.MAD", pivotMad));
CustomFunctions.associate("PIVOT_IQR", protectedFunction("PIVOT.IQR", pivotIqr));
