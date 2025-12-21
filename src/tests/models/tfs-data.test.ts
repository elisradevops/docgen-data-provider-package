import {
  Query,
  QueryTree,
  Column,
  Workitem,
  value,
  QueryType,
  TestCase,
  TestSteps,
  createLinkedRelation,
  createRequirementRelation,
  createMomRelation,
  createBugRelation,
  LinkedRelation,
} from '../../models/tfs-data';

describe('tfs-data models', () => {
  describe('Query class', () => {
    it('should create Query with default arrays', () => {
      const query = new Query();
      expect(query.columns).toEqual([]);
      expect(query.workItems).toEqual([]);
    });
  });

  describe('Column class', () => {
    it('should create Column with properties', () => {
      const column = new Column();
      column.referenceName = 'System.Title';
      column.name = 'Title';
      column.url = 'https://example.com';

      expect(column.referenceName).toBe('System.Title');
      expect(column.name).toBe('Title');
      expect(column.url).toBe('https://example.com');
    });
  });

  describe('Workitem class', () => {
    it('should create Workitem with default Source of 0', () => {
      const workitem = new Workitem();
      expect(workitem.Source).toBe(0);
    });

    it('should allow setting all properties', () => {
      const workitem = new Workitem();
      workitem.id = 123;
      workitem.url = 'https://example.com/workitem/123';
      workitem.parentId = 100;
      workitem.fields = [];
      workitem.attachments = [];
      workitem.level = 1;

      expect(workitem.id).toBe(123);
      expect(workitem.url).toBe('https://example.com/workitem/123');
      expect(workitem.parentId).toBe(100);
      expect(workitem.level).toBe(1);
    });
  });

  describe('value class', () => {
    it('should create value with name and value', () => {
      const v = new value();
      v.name = 'Title';
      v.value = 'Test Value';

      expect(v.name).toBe('Title');
      expect(v.value).toBe('Test Value');
    });
  });

  describe('TestCase class', () => {
    it('should create TestCase with empty relations array', () => {
      const testCase = new TestCase();
      expect(testCase.relations).toEqual([]);
      expect(testCase.caseEvidenceAttachments).toEqual([]);
    });

    it('should allow setting all properties', () => {
      const testCase = new TestCase();
      testCase.id = '123';
      testCase.title = 'Test Case Title';
      testCase.description = 'Test Description';
      testCase.area = 'Area\\Path';
      testCase.steps = [];
      testCase.suit = 'Suite 1';
      testCase.url = 'https://example.com/testcase/123';

      expect(testCase.id).toBe('123');
      expect(testCase.title).toBe('Test Case Title');
      expect(testCase.description).toBe('Test Description');
      expect(testCase.area).toBe('Area\\Path');
      expect(testCase.suit).toBe('Suite 1');
    });
  });

  describe('TestSteps class', () => {
    it('should create TestSteps with all properties', () => {
      const step = new TestSteps();
      step.stepId = '1';
      step.stepPosition = '1';
      step.action = 'Click button';
      step.expected = 'Button clicked';
      step.isSharedStepTitle = false;

      expect(step.stepId).toBe('1');
      expect(step.stepPosition).toBe('1');
      expect(step.action).toBe('Click button');
      expect(step.expected).toBe('Button clicked');
      expect(step.isSharedStepTitle).toBe(false);
    });
  });

  describe('createLinkedRelation', () => {
    it('should create a LinkedRelation object', () => {
      const result = createLinkedRelation(
        '123',
        'Requirement',
        'Test Requirement',
        'https://example.com/wi/123',
        'Parent'
      );

      expect(result).toEqual({
        id: '123',
        wiType: 'Requirement',
        title: 'Test Requirement',
        url: 'https://example.com/wi/123',
        relationType: 'Parent',
      });
    });
  });

  describe('createRequirementRelation', () => {
    it('should create a RequirementRelation with customerId', () => {
      const result = createRequirementRelation('123', 'Requirement Title', 'CUST-001');

      expect(result).toEqual({
        type: 'requirement',
        id: '123',
        title: 'Requirement Title',
        customerId: 'CUST-001',
      });
    });

    it('should create a RequirementRelation without customerId', () => {
      const result = createRequirementRelation('123', 'Requirement Title');

      expect(result).toEqual({
        type: 'requirement',
        id: '123',
        title: 'Requirement Title',
        customerId: undefined,
      });
    });
  });

  describe('createMomRelation', () => {
    it('should create a MomRelation object', () => {
      const result = createMomRelation('123', 'Bug', 'Bug Title', 'https://example.com/wi/123', 'Active');

      expect(result).toEqual({
        type: 'Bug',
        id: '123',
        title: 'Bug Title',
        url: 'https://example.com/wi/123',
        status: 'Active',
      });
    });
  });

  describe('createBugRelation', () => {
    it('should create a BugRelation with severity', () => {
      const result = createBugRelation('123', 'Bug Title', '1 - Critical');

      expect(result).toEqual({
        type: 'bug',
        id: '123',
        title: 'Bug Title',
        severity: '1 - Critical',
      });
    });

    it('should create a BugRelation without severity', () => {
      const result = createBugRelation('123', 'Bug Title');

      expect(result).toEqual({
        type: 'bug',
        id: '123',
        title: 'Bug Title',
        severity: undefined,
      });
    });
  });

  describe('QueryType enum', () => {
    it('should have correct enum values', () => {
      expect(QueryType.Flat).toBe('flat');
      expect(QueryType.Tree).toBe('tree');
      expect(QueryType.OneHop).toBe('oneHop');
    });
  });
});
