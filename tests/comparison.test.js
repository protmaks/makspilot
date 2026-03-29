/**
 * @vitest-environment jsdom
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';

// Mocking dependencies if necessary
// In a real environment, these would be loaded via script tags
// For testing, we might need to mock parts of XLSX or other globals

const {
    prepareDataForComparison,
    createColumnMapping,
    compareValuesWithTolerance
} = require('../javascript/functions.js');

describe('Excel Comparison Logic', () => {
    
    describe('Column Mapping', () => {
        it('should correctly map identical columns', () => {
            const header1 = ['ID', 'Name', 'Date'];
            const header2 = ['ID', 'Name', 'Date'];
            const mapping = createColumnMapping(header1, header2);
            
            expect(mapping.commonColumns).toHaveLength(3);
            expect(mapping.commonColumns[0].name).toBe('ID');
            expect(mapping.onlyInFile1).toHaveLength(0);
            expect(mapping.onlyInFile2).toHaveLength(0);
        });

        it('should handle reordered columns', () => {
            const header1 = ['ID', 'Name', 'Date'];
            const header2 = ['Date', 'ID', 'Name'];
            const mapping = createColumnMapping(header1, header2);
            
            expect(mapping.commonColumns).toHaveLength(3);
            expect(mapping.commonColumns.find(c => c.name === 'Date').index1).toBe(2);
            expect(mapping.commonColumns.find(c => c.name === 'Date').index2).toBe(0);
        });

        it('should identify columns present in only one file', () => {
            const header1 = ['ID', 'Name', 'Extra1'];
            const header2 = ['ID', 'Name', 'Extra2'];
            const mapping = createColumnMapping(header1, header2);
            
            expect(mapping.commonColumns).toHaveLength(2);
            expect(mapping.onlyInFile1).toContain('Extra1');
            expect(mapping.onlyInFile2).toContain('Extra2');
        });

        it('should be case-insensitive when mapping columns', () => {
            const header1 = ['id', 'NAME'];
            const header2 = ['ID', 'name'];
            const mapping = createColumnMapping(header1, header2);
            
            expect(mapping.commonColumns).toHaveLength(2);
        });
    });

    describe('Data Preparation for Comparison', () => {
        it('should align data based on common columns', () => {
            const data1 = [
                ['ID', 'Value'],
                ['1', 'A'],
                ['2', 'B']
            ];
            const data2 = [
                ['Value', 'ID'],
                ['A', '1'],
                ['C', '2']
            ];
            
            const prepared = prepareDataForComparison(data1, data2);
            
            // data1 should have [ID, Value]
            // data2 should now also have [ID, Value] because it's reordered to match data1's common columns
            expect(prepared.data1[0]).toEqual(['ID', 'Value']);
            expect(prepared.data2[0]).toEqual(['ID', 'Value']);
            
            expect(prepared.data1[1]).toEqual(['1', 'A']);
            expect(prepared.data2[1]).toEqual(['1', 'A']);
            expect(prepared.data2[2]).toEqual(['2', 'C']);
        });
    });

    describe('Value Comparison with Tolerance', () => {
        it('should identify identical strings', () => {
            expect(compareValuesWithTolerance('Test', 'Test')).toBe('identical');
            expect(compareValuesWithTolerance('Test', 'test')).toBe('identical'); // case-insensitive
        });

        it('should handle numeric tolerance', () => {
            // Default tolerance is 1.5%
            expect(compareValuesWithTolerance('100', '101')).toBe('tolerance');
            expect(compareValuesWithTolerance('100', '102')).toBe('different');
        });

        it('should handle date tolerance (same date, different time)', () => {
            expect(compareValuesWithTolerance('2023-01-01 10:00', '2023-01-01 15:30')).toBe('tolerance');
            expect(compareValuesWithTolerance('2023-01-01', '2023-01-02')).toBe('different');
        });
    });
});
