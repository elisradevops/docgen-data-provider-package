export default class DataProviderUtils {
  /**
   * Adds a value to an array stored in a Map. If the key doesn't exist,
   * it initializes a new array for that key.
   *
   * @param map The Map to add to
   * @param key The key for the array
   * @param value The value to add to the array
   */
  public static addToTraceMap(map: Map<string, string[]>, key: string, value: string): void {
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key)?.push(value);
  }
}
