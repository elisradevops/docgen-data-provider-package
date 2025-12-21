import { Helper, suiteData, Relations, Links, Trace } from '../../helpers/helper';
import { Query, Workitem } from '../../models/tfs-data';

describe('Helper', () => {
  beforeEach(() => {
    // Reset static state before each test
    Helper.suitList = [];
    Helper.level = 1;
    Helper.first = true;
    Helper.levelList = [];
  });

  describe('suiteData class', () => {
    it('should create suiteData with correct properties', () => {
      // Act
      const suite = new suiteData('Test Suite', '123', '456', 2);

      // Assert
      expect(suite.name).toBe('Test Suite');
      expect(suite.id).toBe('123');
      expect(suite.parent).toBe('456');
      expect(suite.level).toBe(2);
      expect(suite.url).toBeUndefined();
    });
  });

  describe('Relations class', () => {
    it('should create Relations with empty rels array', () => {
      // Act
      const relations = new Relations();

      // Assert
      expect(relations.id).toBeUndefined();
      expect(relations.rels).toEqual([]);
    });

    it('should allow adding relations', () => {
      // Arrange
      const relations = new Relations();
      relations.id = '123';

      // Act
      relations.rels.push('456');
      relations.rels.push('789');

      // Assert
      expect(relations.rels).toEqual(['456', '789']);
    });
  });

  describe('Links class', () => {
    it('should create Links with all properties', () => {
      // Act
      const link = new Links();
      link.id = '123';
      link.title = 'Test Link';
      link.description = 'Test Description';
      link.url = 'https://example.com';
      link.type = 'Parent';
      link.customerId = 'CUST-001';

      // Assert
      expect(link.id).toBe('123');
      expect(link.title).toBe('Test Link');
      expect(link.description).toBe('Test Description');
      expect(link.url).toBe('https://example.com');
      expect(link.type).toBe('Parent');
      expect(link.customerId).toBe('CUST-001');
    });
  });

  describe('Trace class', () => {
    it('should create Trace with all properties', () => {
      // Act
      const trace = new Trace();
      trace.id = '123';
      trace.title = 'Test Trace';
      trace.url = 'https://example.com';
      trace.customerId = 'CUST-001';
      trace.links = [];

      // Assert
      expect(trace.id).toBe('123');
      expect(trace.title).toBe('Test Trace');
      expect(trace.url).toBe('https://example.com');
      expect(trace.customerId).toBe('CUST-001');
      expect(trace.links).toEqual([]);
    });
  });

  describe('findSuitesRecursive', () => {
    const mockUrl = 'https://dev.azure.com/org/';
    const mockProject = 'TestProject';
    const mockPlanId = '100';

    it('should find direct children of a root suite', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite 1', parentSuiteId: '1' },
        { id: '3', title: 'Child Suite 2', parentSuiteId: '1' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result.length).toBe(2);
      expect(result[0].name).toBe('Child Suite 1');
      expect(result[0].id).toBe('2');
      expect(result[0].parent).toBe('1');
      expect(result[1].name).toBe('Child Suite 2');
    });

    it('should find nested suites recursively', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite 1', parentSuiteId: '1' },
        { id: '3', title: 'Grandchild Suite', parentSuiteId: '2' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result.length).toBe(2);
      expect(result[0].name).toBe('Child Suite 1');
      expect(result[1].name).toBe('Grandchild Suite');
    });

    it('should not recurse when recursive is false', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite 1', parentSuiteId: '1' },
        { id: '3', title: 'Grandchild Suite', parentSuiteId: '2' },
      ];

      // Act - When recursive is false and starting from root (parentSuiteId=0),
      // the function returns early without adding children
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', false);

      // Assert - The root suite with parentSuiteId=0 doesn't get added to results
      // and recursive=false means we return immediately
      expect(result.length).toBe(0);
    });

    it('should handle finding a nested suite by ID', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite 1', parentSuiteId: '1' },
        { id: '3', title: 'Grandchild Suite', parentSuiteId: '2' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '2', true);

      // Assert
      expect(result.length).toBe(2);
      expect(result[0].name).toBe('Child Suite 1');
      expect(result[1].name).toBe('Grandchild Suite');
    });

    it('should set correct URLs for suites', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite', parentSuiteId: '1' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result[0].url).toBe(
        `${mockUrl}${mockProject}/_testManagement?planId=${mockPlanId}&suiteId=2&_a=tests`
      );
    });

    it('should return empty array when no children found', () => {
      // Arrange
      const suits = [{ id: '1', title: 'Root Suite', parentSuiteId: 0 }];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result).toEqual([]);
    });

    it('should handle root suite with parentSuiteId = 0', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root Suite', parentSuiteId: 0 },
        { id: '2', title: 'Child Suite', parentSuiteId: '1' },
      ];

      // Act - Find starting from root with recursive=true to get children
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert - Should find the child suite
      expect(result.length).toBe(1);
      expect(result[0].name).toBe('Child Suite');
    });

    it('should handle deeply nested suites', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root', parentSuiteId: 0 },
        { id: '2', title: 'Level 1', parentSuiteId: '1' },
        { id: '3', title: 'Level 2', parentSuiteId: '2' },
        { id: '4', title: 'Level 3', parentSuiteId: '3' },
        { id: '5', title: 'Level 4', parentSuiteId: '4' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result.length).toBe(4);
      expect(result.map((s) => s.name)).toEqual(['Level 1', 'Level 2', 'Level 3', 'Level 4']);
    });

    it('should handle multiple children at same level', () => {
      // Arrange
      const suits = [
        { id: '1', title: 'Root', parentSuiteId: 0 },
        { id: '2', title: 'Child A', parentSuiteId: '1' },
        { id: '3', title: 'Child B', parentSuiteId: '1' },
        { id: '4', title: 'Child C', parentSuiteId: '1' },
        { id: '5', title: 'Grandchild A1', parentSuiteId: '2' },
        { id: '6', title: 'Grandchild B1', parentSuiteId: '3' },
      ];

      // Act
      const result = Helper.findSuitesRecursive(mockPlanId, mockUrl, mockProject, suits, '1', true);

      // Assert
      expect(result.length).toBe(5);
    });
  });

  describe('LevelBuilder', () => {
    it('should build work item hierarchy from query results', () => {
      // Arrange
      const mockQuery: Query = {
        workItems: [
          { Source: 0, fields: [{ value: '1' }], level: 0 } as unknown as Workitem,
          { Source: '1', fields: [{ value: '2' }], level: 0 } as unknown as Workitem,
          { Source: '2', fields: [{ value: '3' }], level: 0 } as unknown as Workitem,
        ],
      } as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '0');

      // Assert
      expect(result.length).toBeGreaterThan(0);
    });

    it('should set level 0 for root items with Source = 0', () => {
      // Arrange
      const mockQuery: Query = {
        workItems: [{ Source: 0, fields: [{ value: '1' }], level: 0 } as unknown as Workitem],
      } as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '0');

      // Assert
      expect(result[0].level).toBe(0);
    });

    it('should build hierarchy with correct levels', () => {
      // Arrange
      const mockQuery: Query = {
        workItems: [
          { Source: 0, fields: [{ value: '1' }], level: 0 } as unknown as Workitem,
          { Source: '1', fields: [{ value: '2' }], level: 0 } as unknown as Workitem,
        ],
      } as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '1');

      // Assert
      expect(result.length).toBe(2);
    });

    it('should handle empty work items array', () => {
      // Arrange
      const mockQuery = {
        workItems: [],
      } as unknown as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '0');

      // Assert
      expect(result).toEqual([]);
    });

    it('should not duplicate work items', () => {
      // Arrange
      const workItem = { Source: 0, fields: [{ value: '1' }], level: 0 } as unknown as Workitem;
      const mockQuery: Query = {
        workItems: [workItem, workItem],
      } as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '0');

      // Assert
      // Should only include the work item once
      expect(result.length).toBe(1);
    });

    it('should handle nested hierarchy', () => {
      // Arrange
      const mockQuery: Query = {
        workItems: [
          { Source: 0, fields: [{ value: '100' }], level: 0 } as unknown as Workitem,
          { Source: '100', fields: [{ value: '200' }], level: 0 } as unknown as Workitem,
          { Source: '200', fields: [{ value: '300' }], level: 0 } as unknown as Workitem,
        ],
      } as Query;

      // Act
      const result = Helper.LevelBuilder(mockQuery, '100');

      // Assert
      expect(result.length).toBe(3);
    });
  });
});
