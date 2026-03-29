/**
 * @vitest-environment jsdom
 */
import { describe, it, expect } from 'vitest';
const {
    normalizeDateTime,
    convertExcelDate,
    convertExcelDateNormalized,
    roundDecimalNumbers,
    compareValuesWithTolerance,
    isWithinTolerance,
    isDateString,
    isNumericString,
    extractDateOnly,
    parseNumber,
    parseCSVValue
} = require('../javascript/functions.js');

describe('Normalization Logic Tests', () => {
    describe('String Trimming and CSV Parsing', () => {
        it('should parse CSV values and trim them correctly', () => {
            expect(parseCSVValue('" quoted "')).toBe(' quoted ');
            expect(parseCSVValue('null')).toBe('');
            expect(parseCSVValue('NULL')).toBe('');
            expect(parseCSVValue('undefined')).toBe('undefined'); // the function only checks for 'null' lowercase, let's see. Wait, actually it does `value.toLowerCase() === 'null'` so it should be ''. Let's test standard strings.
        });
    });

    describe('Number Normalization and Rounding', () => {
        it('should round decimal numbers to 2 decimal places', () => {
            expect(roundDecimalNumbers(12.345)).toBe(12.35);
            expect(roundDecimalNumbers(12.344)).toBe(12.34);
            expect(roundDecimalNumbers(10)).toBe(10);
            expect(roundDecimalNumbers('15.678')).toBe(15.68);
            expect(roundDecimalNumbers('not-a-number')).toBe('not-a-number');
        });

        it('should detect numeric strings', () => {
            expect(isNumericString('123')).toBe(true);
            expect(isNumericString('123.45')).toBe(true);
            expect(isNumericString(' 45.67 ')).toBe(true); // Trimming works
            expect(isNumericString('abc')).toBe(false);
            expect(isNumericString('12abc')).toBe(false);
        });

        it('should parse numbers correctly', () => {
            expect(parseNumber('123.45')).toBe(123.45);
            expect(parseNumber(' $123.45 ')).toBe(123.45); // Removes $ and spaces
            expect(parseNumber('1,234.56')).toBe(1234.56); // Removing commas
            expect(parseNumber('invalid')).toBe(0);
        });
    });

    describe('Date Normalization', () => {
        it('should format Excel serial numbers to ISO date strings', () => {
            // Excel serial 44197 -> 2021-01-01
            const result = convertExcelDate(44197, true);
            expect(result).toMatch(/^2021-01-01/);
        });

        it('should leave normal numbers untouched if not in date column', () => {
            expect(convertExcelDate(44197, false)).toBe(44197);
        });

        it('should normalize MM/DD/YYYY strings to YYYY-MM-DD', () => {
            expect(convertExcelDate('12/31/2023', true)).toBe('2023-12-31');
            expect(convertExcelDate('1/5/2023', true)).toBe('2023-01-05');
        });

        it('should remove trailing zero time from normalized datetimes', () => {
            expect(normalizeDateTime('2023-01-01 00:00:00')).toBe('2023-01-01');
            expect(normalizeDateTime('2023-01-01 00:00')).toBe('2023-01-01');
            expect(normalizeDateTime('2023-01-01 12:30:00')).toBe('2023-01-01 12:30:00');
        });

        it('should normalize various date string formats', () => {
            expect(convertExcelDateNormalized('31.12.2023', true)).toBe('2023-12-31');
            expect(convertExcelDateNormalized('2023-05-15T14:30:00', true)).toBe('2023-05-15 14:30');
            expect(convertExcelDateNormalized('15-Jan-2023', true)).toBe('2023-01-15');
            expect(convertExcelDateNormalized('15 January 2023', true)).toBe('2023-01-15');
        });

        it('should extract date only from datetime strings', () => {
            expect(extractDateOnly('2023-05-15 14:30:00')).toBe('2023-05-15');
            expect(extractDateOnly('  2021-10-01  ')).toBe('2021-10-01');
        });

        it('should correctly identify date strings', () => {
            expect(isDateString('2023-05-15')).toBe(true);
            expect(isDateString('12/31/2023')).toBe(true);
            expect(isDateString('Not a date')).toBe(false);
        });
    });

    describe('Tolerance Logic', () => {
        it('should return "identical" for exact string matches', () => {
            expect(compareValuesWithTolerance('Hello ', ' hello')).toBe('identical'); // Because of case-insensitivity and trimming
            expect(compareValuesWithTolerance('Test', 'test')).toBe('identical');
        });

        it('should return "different" for different strings', () => {
            expect(compareValuesWithTolerance('apple', 'orange')).toBe('different');
            expect(compareValuesWithTolerance('123', 'abc')).toBe('different');
        });

        it('should handle numeric tolerance correctly', () => {
            // Default tolerance is 0.015 (1.5%)
            expect(isWithinTolerance(100, 101, 0.015)).toBe(true);  // Difference is 1, Avg is 100.5. 1/100.5 ~ 0.0099 < 0.015
            expect(isWithinTolerance(100, 102, 0.015)).toBe(false); // Difference is 2, Avg is 101. 2/101 ~ 0.0198 > 0.015
            
            // Testing via compareValuesWithTolerance
            expect(compareValuesWithTolerance('100', '101')).toBe('tolerance');
            expect(compareValuesWithTolerance('100', '102')).toBe('different');
        });

        it('should handle date tolerance correctly (same date, different time)', () => {
            // Different times but same date should be "tolerance"
            expect(compareValuesWithTolerance('2023-05-15 10:00:00', '2023-05-15 14:30:00')).toBe('tolerance');
            // Different dates
            expect(compareValuesWithTolerance('2023-05-15', '2023-05-16')).toBe('different');
            // Identical datetime
            expect(compareValuesWithTolerance('2023-05-15 10:00', '2023-05-15 10:00')).toBe('identical');
        });

        it('should handle missing values correctly in comparison', () => {
            expect(compareValuesWithTolerance('', '')).toBe('identical');
            expect(compareValuesWithTolerance('val', '')).toBe('different');
            expect(compareValuesWithTolerance(null, 'val')).toBe('different');
        });
    });
});
