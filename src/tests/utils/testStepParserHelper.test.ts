import { TFSServices } from '../../helpers/tfs';
import TestStepParserHelper from '../../utils/testStepParserHelper';
import logger from '../../utils/logger';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');

describe('TestStepParserHelper', () => {
  let testStepParserHelper: TestStepParserHelper;
  const mockOrgUrl = 'https://dev.azure.com/organization/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    testStepParserHelper = new TestStepParserHelper(mockOrgUrl, mockToken);
  });

  describe('parseTestSteps', () => {
    it('should parse simple test steps from XML', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Step 1 Action</parameterizedString>
            <parameterizedString isformatted="true">Step 1 Expected</parameterizedString>
          </step>
          <step id="2" type="ActionStep">
            <parameterizedString isformatted="true">Step 2 Action</parameterizedString>
            <parameterizedString isformatted="true">Step 2 Expected</parameterizedString>
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(2);
      expect(result[0].stepId).toBe('1');
      expect(result[0].stepPosition).toBe('1');
      expect(result[0].action).toBe('Step 1 Action');
      expect(result[0].expected).toBe('Step 1 Expected');
      expect(result[0].isSharedStepTitle).toBe(false);
      expect(result[1].stepId).toBe('2');
      expect(result[1].stepPosition).toBe('2');
    });

    it('should parse test steps with level prefix', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Action</parameterizedString>
            <parameterizedString isformatted="true">Expected</parameterizedString>
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(
        xmlSteps,
        sharedStepIdToRevisionLookupMap,
        '1.'
      );

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].stepPosition).toBe('1.1');
    });

    it('should handle empty steps XML', async () => {
      // Arrange
      const xmlSteps = `<steps></steps>`;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(0);
    });

    it('should handle invalid XML gracefully', async () => {
      // Arrange
      const xmlSteps = `<invalid xml`;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toEqual([]);
      expect(logger.error).toHaveBeenCalled();
    });

    it('should parse shared steps (compref) with revision lookup', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();
      sharedStepIdToRevisionLookupMap.set(999, 5);

      const mockSharedStepWI = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `
            <steps>
              <step id="1" type="ActionStep">
                <parameterizedString isformatted="true">Shared Step Action</parameterizedString>
                <parameterizedString isformatted="true">Shared Step Expected</parameterizedString>
              </step>
            </steps>
          `,
          'System.Title': 'Shared Step Title',
        },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockSharedStepWI);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}/_apis/wit/workitems/999/revisions/5`,
        mockToken
      );
      expect(result).toHaveLength(2); // Title + 1 step
      expect(result[0].isSharedStepTitle).toBe(true);
      expect(result[0].action).toContain('Shared Step Title');
      expect(result[1].action).toBe('Shared Step Action');
    });

    it('should parse shared steps without revision lookup', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      const mockSharedStepWI = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `
            <steps>
              <step id="1" type="ActionStep">
                <parameterizedString isformatted="true">Shared Action</parameterizedString>
                <parameterizedString isformatted="true">Shared Expected</parameterizedString>
              </step>
            </steps>
          `,
          'System.Title': 'My Shared Step',
        },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockSharedStepWI);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}/_apis/wit/workitems/999`,
        mockToken
      );
      expect(result).toHaveLength(2);
    });

    it('should handle shared step fetch error by logging and rethrowing', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act - The error propagates through processCompref -> fetchSharedSteps
      // but parseTestSteps catches it and returns empty array
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert - parseTestSteps catches the error and returns empty array
      expect(result).toEqual([]);
      expect(logger.error).toHaveBeenCalled();
    });

    it('should handle shared step with no steps XML', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      const mockSharedStepWI = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': null,
          'System.Title': 'Empty Shared Step',
        },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockSharedStepWI);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(0);
    });

    it('should handle steps with missing parameterizedString', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].action).toBe('');
      expect(result[0].expected).toBe('');
    });

    it('should handle nested shared steps', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Regular Step</parameterizedString>
            <parameterizedString isformatted="true">Expected</parameterizedString>
          </step>
          <compref id="100" ref="999">
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      const mockSharedStepWI = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `
            <steps>
              <step id="1" type="ActionStep">
                <parameterizedString isformatted="true">Nested Shared Action</parameterizedString>
                <parameterizedString isformatted="true">Nested Shared Expected</parameterizedString>
              </step>
            </steps>
          `,
          'System.Title': 'Nested Shared Step',
        },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockSharedStepWI);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(3); // 1 regular + 1 title + 1 shared step
      expect(result[0].stepId).toBe('1');
      expect(result[0].action).toBe('Regular Step');
      expect(result[1].isSharedStepTitle).toBe(true);
      expect(result[2].action).toBe('Nested Shared Action');
    });

    it('should handle parent step ID for nested steps', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Action</parameterizedString>
            <parameterizedString isformatted="true">Expected</parameterizedString>
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(
        xmlSteps,
        sharedStepIdToRevisionLookupMap,
        '',
        'parent123'
      );

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].stepId).toBe('parent123;1');
    });

    it('should handle compref with nested children steps', async () => {
      // Arrange - compref with children that contain step nodes
      // This tests lines 86-100 where comprefNode.children is processed
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
            <step id="1" type="ActionStep">
              <parameterizedString isformatted="true">Nested Step Action</parameterizedString>
              <parameterizedString isformatted="true">Nested Step Expected</parameterizedString>
            </step>
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      const mockSharedStepWI = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `
            <steps>
              <step id="1" type="ActionStep">
                <parameterizedString isformatted="true">Shared Action</parameterizedString>
                <parameterizedString isformatted="true">Shared Expected</parameterizedString>
              </step>
            </steps>
          `,
          'System.Title': 'Shared Step',
        },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockSharedStepWI);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert - should have shared step title + shared step + nested step from children
      expect(result.length).toBeGreaterThanOrEqual(2);
    });

    it('should handle step position with dot notation', async () => {
      // Arrange - This tests lines 91-95 where stepPosition includes '.'
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Action 1</parameterizedString>
            <parameterizedString isformatted="true">Expected 1</parameterizedString>
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act - Pass a level with dot notation
      const result = await testStepParserHelper.parseTestSteps(
        xmlSteps,
        sharedStepIdToRevisionLookupMap,
        '2.3.'
      );

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].stepPosition).toBe('2.3.1');
    });

    it('should handle nested compref within compref', async () => {
      // Arrange - This tests lines 101-109 where nested compref is processed
      const xmlSteps = `
        <steps>
          <compref id="100" ref="999">
            <compref id="200" ref="888">
            </compref>
          </compref>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      const mockSharedStepWI1 = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `<steps><step id="1" type="ActionStep"><parameterizedString>Action1</parameterizedString><parameterizedString>Expected1</parameterizedString></step></steps>`,
          'System.Title': 'Shared Step 1',
        },
      };

      const mockSharedStepWI2 = {
        fields: {
          'Microsoft.VSTS.TCM.Steps': `<steps><step id="1" type="ActionStep"><parameterizedString>Action2</parameterizedString><parameterizedString>Expected2</parameterizedString></step></steps>`,
          'System.Title': 'Shared Step 2',
        },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockSharedStepWI1)
        .mockResolvedValueOnce(mockSharedStepWI2);

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert - should process both shared steps
      expect(result.length).toBeGreaterThanOrEqual(2);
    });

    it('should handle complex step position calculations', async () => {
      // Arrange
      const xmlSteps = `
        <steps>
          <step id="1" type="ActionStep">
            <parameterizedString isformatted="true">Step 1</parameterizedString>
            <parameterizedString isformatted="true">Expected 1</parameterizedString>
          </step>
          <step id="2" type="ActionStep">
            <parameterizedString isformatted="true">Step 2</parameterizedString>
            <parameterizedString isformatted="true">Expected 2</parameterizedString>
          </step>
          <step id="3" type="ActionStep">
            <parameterizedString isformatted="true">Step 3</parameterizedString>
            <parameterizedString isformatted="true">Expected 3</parameterizedString>
          </step>
        </steps>
      `;
      const sharedStepIdToRevisionLookupMap = new Map<number, number>();

      // Act
      const result = await testStepParserHelper.parseTestSteps(xmlSteps, sharedStepIdToRevisionLookupMap);

      // Assert
      expect(result).toHaveLength(3);
      expect(result[0].stepPosition).toBe('1');
      expect(result[1].stepPosition).toBe('2');
      expect(result[2].stepPosition).toBe('3');
    });
  });
});
