import DataProviderUtils from '../../utils/DataProviderUtils';

describe('DataProviderUtils', () => {
  describe('addToTraceMap', () => {
    it('should add value to new key in map', () => {
      // Arrange
      const map = new Map<string, string[]>();

      // Act
      DataProviderUtils.addToTraceMap(map, 'key1', 'value1');

      // Assert
      expect(map.has('key1')).toBe(true);
      expect(map.get('key1')).toEqual(['value1']);
    });

    it('should add value to existing key in map', () => {
      // Arrange
      const map = new Map<string, string[]>();
      map.set('key1', ['value1']);

      // Act
      DataProviderUtils.addToTraceMap(map, 'key1', 'value2');

      // Assert
      expect(map.get('key1')).toEqual(['value1', 'value2']);
    });

    it('should handle multiple keys', () => {
      // Arrange
      const map = new Map<string, string[]>();

      // Act
      DataProviderUtils.addToTraceMap(map, 'key1', 'value1');
      DataProviderUtils.addToTraceMap(map, 'key2', 'value2');
      DataProviderUtils.addToTraceMap(map, 'key1', 'value3');

      // Assert
      expect(map.size).toBe(2);
      expect(map.get('key1')).toEqual(['value1', 'value3']);
      expect(map.get('key2')).toEqual(['value2']);
    });

    it('should handle empty string key', () => {
      // Arrange
      const map = new Map<string, string[]>();

      // Act
      DataProviderUtils.addToTraceMap(map, '', 'value1');

      // Assert
      expect(map.has('')).toBe(true);
      expect(map.get('')).toEqual(['value1']);
    });

    it('should handle empty string value', () => {
      // Arrange
      const map = new Map<string, string[]>();

      // Act
      DataProviderUtils.addToTraceMap(map, 'key1', '');

      // Assert
      expect(map.get('key1')).toEqual(['']);
    });

    it('should not throw when key exists but stored value is undefined', () => {
      const map = new Map<string, string[]>();
      map.set('key1', undefined as any);

      DataProviderUtils.addToTraceMap(map, 'key1', 'value1');

      expect(map.get('key1')).toBeUndefined();
    });
  });
});
